import { LEGACY_SYNC_API_PROP } from '../components/DocxEditor/peerless/contracts';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { StrictMode, createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import type { AsyncReplicaAdapter, CollaborationReplica } from '@betteroffice/docx/collaboration';
import { encodeSyncStep2, encodeUpdate } from '@betteroffice/docx/collaboration/protocol';
import { createResidentEngineSession } from '@betteroffice/docx/yrs/residentEngineSession';
import {
  createYrsSession,
  takePreloadedResidentEngineWorker,
  withYrsSelectionCause,
  type DocxEditRequest,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import * as yrs from '@betteroffice/docx/yrs';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
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
  type DocxEditorBaseProps,
  type DocxEditorPluginProps,
  type DocxEditorRef,
  type DocxPlugin,
  type DocxPluginCommandResult,
  type DocxPluginContext,
  type DocxPluginDefinition,
  type DocxPluginError,
  type DocxPluginEvent,
  type DocxPluginGeometry,
  type DocxPluginSelection,
  type DocxAnchorGeometryResult,
  type DocxAnchorRect,
  type DocxGeometryTarget,
} from '../index';
import { isMacPlatform } from '../commands/descriptors';
import * as pluginHosts from './createDocxPluginHost';
import * as workerPluginGeometry from '../components/DocxEditor/peerless/workerPluginGeometry';
import * as shellControllers from '../components/DocxEditor/peerless/shellController';
import * as pagePointers from '../components/DocxEditor/hooks/usePagesPointer';
import * as docxBindings from './useDocxPluginHost';
import { selectionClock, selectionWindow } from '../../../../e2e/browser/plugin-host/selection-clock';
import { storyAnchorsDocx } from './__fixtures__/storyAnchorsDocx';
import { scopePluginGeometry } from './geometry';
import * as canvasReplay from '../components/DocxEditor/canvasReplay';
import * as headerFooterEditing from '../components/DocxEditor/hooks/useHeaderFooterEditing';
import {
  resetEngineChoiceForTests,
  setMissingWorkerCapabilitiesForTests,
} from '../components/DocxEditor/internals/engineChoice';
import { setupWorkerEngine } from '../components/DocxEditor/__fixtures__/workerEngine';
import { testScope } from '../components/DocxEditor/__fixtures__/testScope';

const mount = (props: Partial<DocxEditorBaseProps> = {}) => mountEditor(props, false, documentBytes(), false);

const MOD = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');
const { settle, until } = testScope(act);

const FIXTURE = resolve(
  import.meta.dir,
  '../components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx'
);
const quiet = { error: console.error, warn: console.warn };
const originalWorker = globalThis.Worker;
const workers: InProcessResidentWorker[] = [];
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>> | null = null;
const editorEngines: Array<[string, false | undefined]> = [
  ['in-thread', false],
  ['default worker', undefined],
];

beforeAll(async () => {
  configureDefaultFonts({ fonts: { createFontProvider: () => ({ resolve: () => undefined }) } });
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
  startWorker = await residentWorkerFactory();
  console.error = () => {};
  console.warn = () => {};
}, 30_000);
afterEach(async () => {
  cleanup();
  takePreloadedResidentEngineWorker()?.destroy();
  await act(async () => {});
  compileModule?.mockRestore();
  compileModule = null;
  for (const worker of workers.splice(0)) worker.terminate();
  resetEngineChoiceForTests();
  globalThis.Worker = originalWorker;
});
afterAll(async () => {
  configureDefaultFonts({});
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

async function mountEditor(
  props: Partial<DocxEditorBaseProps>,
  strict: boolean,
  buffer: ArrayBuffer,
  experimentalWorkerOpen: false | undefined
) {
  const ref = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  const opens = workers.reduce(
    (count, worker) => count + worker.requests.filter((type) => type === 'open').length, 0
  );
  let currentProps = props;
  const onError = (error: Error) => { errors.push(error); currentProps.onError?.(error); };
  const element = (next: Partial<DocxEditorBaseProps>) => {
    currentProps = next;
    const editor = experimentalWorkerOpen === false
      ? <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }} ref={ref} experimentalWorkerOpen={false} documentBuffer={buffer} {...next} onError={onError} />
      : <DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }} ref={ref} documentBuffer={buffer} {...next} onError={onError} />;
    return strict ? <StrictMode>{editor}</StrictMode> : editor;
  };
  const view = render(element(props));
  await until(() => ref.current?.commands.getState('save').enabled === true || errors.length > 0);
  expect(errors).toEqual([]);
  expect(ref.current?.commands.getState('save').enabled).toBe(true);
  if (experimentalWorkerOpen === undefined) {
    expect(workers.length).toBeGreaterThan(0);
    expect(workers.some(
      (worker) => worker.sessions.length > 0 && worker.requests.includes('open')
    )).toBe(true);
    expect(workers.reduce(
      (count, worker) => count + worker.requests.filter((type) => type === 'open').length, 0
    )).toBeGreaterThan(opens);
    await act(async () => {
      await ref.current!.whenLayoutComplete({ timeoutMs: 3000 });
    });
  }
  return { ref, view, rerender: (next: Partial<DocxEditorBaseProps>) => view.rerender(element(next)) };
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

async function selectInThreadFirstWord(ref: React.RefObject<DocxEditorRef | null>) {
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

function installWorkerEngine() {
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  setMissingWorkerCapabilitiesForTests([]);
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
}

type Mount = (props: DocxEditorPluginProps) => Promise<{ ref: React.RefObject<DocxEditorRef | null>; view: ReturnType<typeof render> }>;

async function loadHookAppliesOnce(mount: Mount) {
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
}

describe.each(editorEngines)('DocxEditor plugins (%s engine)', (_engine, experimentalWorkerOpen) => {
  beforeEach(() => {
    if (experimentalWorkerOpen !== false) installWorkerEngine();
  });

  const mount = (
    props: Partial<DocxEditorBaseProps> = {},
    strict = false,
    buffer: ArrayBuffer = documentBytes()
  ) => mountEditor(props, strict, buffer, experimentalWorkerOpen);

  test('selection changes emit one selection-settle after the quiet window', async () => {
    const events: DocxPluginEvent[] = [];
    const plugin = defineDocxPlugin({ id: 'settle', createState: () => null,
      onEvent(_context, event) { events.push(event); } });
    const { ref } = await mount({ plugins: [plugin] });
    await settle(400);
    events.length = 0;
    await act(async () => { ref.current!.getEditorRef()!.setSelection(1, 4); });
    expect(events.filter((event) => event.type === 'selection-settle')).toEqual([]);
    await settle(400);
    expect(events.filter((event) => event.type === 'selection-settle')).toMatchObject([
      { source: 'keyboard', selection: { displayRange: { from: 1, to: 4 } } },
    ]);
    expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
    await settle(400);
    expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
  });

  if (experimentalWorkerOpen === false) {
    test.each(['header', 'footer'] as const)('caret moves in uniformly formatted %s text settle independently of toolbar updates', async (part) => {
      const events: DocxPluginEvent[] = [];
      let context: DocxPluginContext<null> | undefined;
      const onSelectionChange = mock(() => {});
      const plugin = defineDocxPlugin({ id: 'part-settle', createState: () => null,
        initialize(next) { context = next; },
        onEvent(next, event) { context = next; events.push(event); } });
      const original = headerFooterEditing.useHeaderFooterEditing;
      let open: ReturnType<typeof original>['handleHeaderFooterDoubleClick'] | undefined;
      const editing = spyOn(headerFooterEditing, 'useHeaderFooterEditing').mockImplementation((options) => {
        const result = original(options);
        open = result.handleHeaderFooterDoubleClick;
        return result;
      });
      const clock = selectionClock();
      try {
        const { ref } = await mount({ plugins: [plugin], onSelectionChange });
        await act(async () => { open!(part, 1); });
        await until(() => document.activeElement instanceof HTMLTextAreaElement &&
          ref.current!.getEditorRef()!.getYrsSession()!.selection()!.head.story.startsWith('hf:'));
        const input = document.activeElement as HTMLTextAreaElement;
        await act(async () => {
          fireEvent.keyDown(input, { key: 'Home' });
          fireEvent.keyDown(input, { key: 'ArrowRight' });
          fireEvent.keyDown(input, { key: 'ArrowRight' });
          await ref.current!.flushPendingInput();
        });
        expect(ref.current!.getEditorRef()!.getYrsSession()!.selection()!.head.offset).toBe(2);
        await act(async () => { await clock.tick(400); });
        expect(context!.snapshot.selectionSettled).toBe(true);
        const story = ref.current!.getEditorRef()!.getYrsSession()!.selection()!.head.story;
        expect(story.startsWith('hf:')).toBe(true);
        events.length = 0;
        onSelectionChange.mockClear();
        const previousSelection = JSON.stringify(context!.snapshot.selection);
        await act(async () => { fireEvent.keyDown(input, { key: 'ArrowRight' }); await ref.current!.flushPendingInput(); });
        expect(ref.current!.getEditorRef()!.getYrsSession()!.selection()!.head).toMatchObject({ story, offset: 3 });
        expect(onSelectionChange).not.toHaveBeenCalled();
        await until(() => JSON.stringify(context!.snapshot.selection) !== previousSelection);
        expect(context!.snapshot.selectionSettled).toBe(false);
        expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(0);
        await act(async () => { await clock.tick(400); });
        expect(events.filter((event) => event.type === 'selection-settle')).toMatchObject([
          { source: 'keyboard', selection: { displayRange: { story, from: 3, to: 3 } } },
        ]);
        expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
        expect(context!.snapshot.selectionSettled).toBe(true);
        await act(async () => { await clock.tick(400); });
        expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
      } finally { editing.mockRestore(); clock.restore(); }
    });
  }

  test('plugins receive the resolved locale and one event when it changes', async () => {
    const contexts: DocxPluginContext<null>[] = [];
    const events: DocxPluginEvent[] = [];
    const plugin = defineDocxPlugin({ id: 'locale', createState: () => null,
      onEvent(context, event) { contexts.push(context); events.push(event); } });
    const { rerender } = await mount({ plugins: [plugin], i18n: { _lang: 'de' } });
    await until(() => events.some((event) => event.type === 'load'));
    expect(contexts[0]!.snapshot.locale).toBe('de');
    expect(events.filter((event) => event.type === 'locale-change')).toEqual([]);
    rerender({ plugins: [plugin], i18n: { _lang: 'fr' } });
    await until(() => events.some((event) => event.type === 'locale-change'));
    expect(events.filter((event) => event.type === 'locale-change')).toEqual([
      { type: 'locale-change', generation: contexts.at(-1)!.snapshot.generation, locale: 'fr' },
    ]);
    expect(contexts.at(-1)!.snapshot.locale).toBe('fr');
    rerender({ plugins: [plugin], i18n: { _lang: 'fr' } });
    await settle();
    expect(events.filter((event) => event.type === 'locale-change')).toHaveLength(1);
  });

  test('plugin locales are canonical and equivalent tags do not notify twice', async () => {
    const contexts: DocxPluginContext<null>[] = [];
    const events: DocxPluginEvent[] = [];
    const plugin = defineDocxPlugin({ id: 'canonical-locale', createState: () => null,
      onEvent(context, event) { contexts.push(context); events.push(event); } });
    const { rerender } = await mount({ plugins: [plugin], i18n: { _lang: 'DE-de' } });
    await until(() => events.some((event) => event.type === 'load'));
    expect(contexts[0]!.snapshot.locale).toBe('de-DE');
    rerender({ plugins: [plugin], i18n: { _lang: 'de-DE' } });
    await settle();
    expect(events.filter((event) => event.type === 'locale-change')).toEqual([]);
    rerender({ plugins: [plugin], i18n: { _lang: 'invalid_tag' } });
    await until(() => events.some((event) => event.type === 'locale-change'));
    expect(events.filter((event) => event.type === 'locale-change')).toEqual([
      { type: 'locale-change', generation: contexts.at(-1)!.snapshot.generation, locale: 'en' },
    ]);
    expect(contexts.at(-1)!.snapshot.locale).toBe('en');
  });

  for (const readOnly of [false, true]) {
    test(`the portaled document input belongs to the focus roots (readOnly=${readOnly})`, async () => {
      const events: DocxPluginEvent[] = [];
      const snapshots: boolean[] = [];
      const plugin = defineDocxPlugin({ id: 'focus-input', createState: () => null,
        panel: { title: 'Focus', placement: 'right', render: () => <input data-testid="focus-panel" /> },
        onEvent(context, event) {
          if (event.type === 'focus-change') {
            events.push(event);
            snapshots.push(context.snapshot.focused);
          }
        } });
      const { ref, view } = await mount({ plugins: [plugin], readOnly });
      const surface = view.container.querySelector('.paged-editor')!;
      const input = document.getElementById(surface.getAttribute('data-docx-input')!);
      expect(input).toBeInstanceOf(HTMLTextAreaElement);
      expect(surface.contains(input)).toBe(false);
      const panel = view.getByTestId('focus-panel');
      await act(async () => { panel.focus(); });
      await settle();
      events.length = 0;
      snapshots.length = 0;
      await act(async () => { ref.current!.focus(); });
      expect(document.activeElement).toBe(input);
      await until(() => events.some((event) => event.type === 'focus-change' && event.focused));
      await act(async () => { panel.focus(); });
      await until(() => events.some((event) => event.type === 'focus-change' && !event.focused));
      expect(snapshots).toEqual([true, false]);
    });
  }

  const selectFirstWord = async (ref: React.RefObject<DocxEditorRef | null>) => {
    if (experimentalWorkerOpen === false) return selectInThreadFirstWord(ref);
    const { paragraph } = await firstParagraph(ref);
    await act(async () => {
      ref.current!.highlightRange(1, 4);
      await ref.current!.flushPendingInput();
    });
    expect(await ref.current!.readSelectionInfo()).toMatchObject({
      paraId: paragraph.paraId,
      selectedText: paragraph.text.slice(0, 3),
    });
  };

  test('initialize, load, one change per batch, and cleanup on removal and readdition', async () => {
    const { plugin, log, contexts } = recorder();
    let live: CollaborationReplica | null = null;
    let host!: pluginHosts.DocxPluginHost;
    const createHost = pluginHosts.createDocxPluginHost;
    const capture = spyOn(pluginHosts, 'createDocxPluginHost').mockImplementation((access) => {
      host = createHost(access);
      return host;
    });
    let mounted!: Awaited<ReturnType<typeof mount>>;
    try { mounted = await mount({
      plugins: [plugin],
      ...(experimentalWorkerOpen === undefined ? {
        collaboration: { onReplica: (replica: CollaborationReplica | null) => { live = replica; } },
      } : {}),
    }); } finally { capture.mockRestore(); }
    const { ref, rerender } = mounted;
    const drainNotifications = async () => {
      await ref.current!.flushPendingInput();
      await ref.current!.readParagraphs({ view: 'accepted', paraIds: [] });
      await host.notificationsIdle();
    };
    await until(() => log.includes('load:loaded'));
    expect(log.slice(0, 2)).toEqual(['initialize', 'load:loaded']);
    const changeCount = () => log.filter((entry) => entry.startsWith('document-change')).length;
    const loadVersion = contexts[1].snapshot.version;

    const { version, paragraph } = await firstParagraph(ref);
    expect(version).toBe(loadVersion);
    let applied!: Awaited<ReturnType<DocxEditorRef['applyEdits']>>;
    await act(async () => {
      applied = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
    });
    expect(applied).toMatchObject({ ok: true, applied: true });
    await until(() => changeCount() === 1);
    const changes = log.filter((entry) => entry.startsWith('document-change'));
    expect(changes).toEqual([`document-change:${applied.ok ? applied.version : ''}`]);

    await act(async () => {
      const stale = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
      expect(stale).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    });
    await act(drainNotifications);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toHaveLength(1);

    const session = experimentalWorkerOpen === false
      ? ref.current!.getEditorRef()!.getYrsSession()!
      : null;
    for (const step of ['undo', 'redo'] as const) {
      await act(async () => {
        if (session) ref.current!.getEditorRef()![step]();
        else {
          expect(await ref.current!.commands.execute(step, null)).toMatchObject({
            ok: true, status: 'executed',
          });
        }
      });
      await until(() => changeCount() === (step === 'undo' ? 2 : 3));
      expect(log.filter((entry) => entry.startsWith('document-change')).at(-1)).toBe(
        `document-change:${session ? session.version() : (await firstParagraph(ref)).version}`
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
    await act(drainNotifications);
    expect(log.filter((entry) => entry.startsWith('document-change'))).toHaveLength(3);
    const replica = await createYrsSession({ clientId: 5151 });
    try {
      if (!session) {
        await until(() => live !== null);
        expect(live).not.toBeNull();
      }
      const transport: CollaborationReplica = session ?? live!;
      replica.applyUpdate(transport.encodeStateAsUpdate());
      const last = replica.paragraphs('body').at(-1)!;
      replica.insertText(
        { story: 'body', paraId: last.paraId, offset: last.text.length },
        ' remote'
      );
      await act(async () => {
        transport.applyUpdate(replica.encodeStateAsUpdate(transport.encodeStateVector()));
      });
      await until(() => changeCount() === 4);
    } finally {
      replica.destroy();
    }
    const versions = log.filter((entry) => entry.startsWith('document-change'));
    expect(versions).toHaveLength(4);
    expect(versions.at(-1)).toBe(
      `document-change:${session ? session.version() : (await firstParagraph(ref)).version}`
    );

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
  }, 30_000);

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
    const uses = [
      { sectionIndex: 0, variant: 'default' as const },
      { sectionIndex: 1, variant: 'default' as const },
    ];
    const listed = await context.read.listStories();
    expect(listed.ok && listed.stories.filter((story) => story.uses)).toEqual([
      { story: 'hf:rId3', kind: 'header', root: 'hf:rId3', part: expect.stringContaining('header'), uses },
      { story: 'hf:rId4', kind: 'footer', root: 'hf:rId4', part: expect.stringContaining('footer'), uses },
    ]);
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
    const editor = experimentalWorkerOpen === false ? ref.current!.getEditorRef()! : null;
    await act(async () => {
      if (editor) {
        const session = editor.getYrsSession()!;
        session.setSelection({ story: 'body', paraId: paragraph.paraId, offset: original.length });
        editor.syncYrsInputState(false);
        editor.insertText(' typed');
      } else {
        expect(await ref.current!.scrollToParagraph(paragraph.paraId)).toBe(true);
        const input = within(document.body).getByLabelText('Document input');
        ref.current!.focus();
        ref.current!.highlightRange(original.length + 1, original.length + 1);
        await ref.current!.flushPendingInput();
        fireEvent.input(input, {
          target: { value: ' typed' }, inputType: 'insertText', data: ' typed',
        });
        await ref.current!.flushPendingInput();
      }
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
      if (editor) expect(editor.undo()).toBe(true);
      else {
        expect(await ref.current!.commands.execute('undo', null)).toMatchObject({
          ok: true, status: 'executed',
        });
      }
    });
    expect((await firstParagraph(ref)).paragraph.text).toBe(typed);
    await act(async () => {
      if (editor) expect(editor.undo()).toBe(true);
      else {
        expect(await ref.current!.commands.execute('undo', null)).toMatchObject({
          ok: true, status: 'executed',
        });
      }
    });
    expect((await firstParagraph(ref)).paragraph.text).toBe(original);
    await act(async () => {
      if (editor) {
        editor.redo();
        editor.redo();
      } else {
        expect(await ref.current!.commands.execute('redo', null)).toMatchObject({
          ok: true, status: 'executed',
        });
        expect(await ref.current!.commands.execute('redo', null)).toMatchObject({
          ok: true, status: 'executed',
        });
      }
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

  test('a load hook that applies a batch every run loads once and keeps its state', () =>
    loadHookAppliesOnce(mount));

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
            data-drawn={geometry.toOverlayRect({ x: 0, y: 0, width: 1, height: 1 }) !== null}
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

    const mounted = marker();
    const heldEdit = async (
      from: string,
      superseded: readonly DocxPluginGeometry[]
    ): Promise<string> => {
      let applied = '';
      const frames: FrameRequestCallback[] = [];
      const frame = spyOn(globalThis, 'requestAnimationFrame').mockImplementation((callback) => {
        frames.push(callback);
        return frames.length;
      });
      try {
        await act(async () => {
          const result = await ref.current!.applyEdits(appendRequest(from, paragraph.paraId));
          if (result.ok) applied = result.version;
        });
        expect(applied).not.toBe('');
        expect(marker()).toBe(mounted);
        expect(marker()!.dataset).toMatchObject({
          version: from,
          snapshot: applied,
          drawn: 'true',
        });
        expect(geometries.at(-1)!.layout.version).toBe(from);
        expect(geometries.at(-1)!.toOverlayRect(unit)).not.toBeNull();
        expect(contexts.at(-1)!.geometry).toBeNull();
        for (const old of superseded) expect(old.toOverlayRect(unit)).toBeNull();
      } finally {
        frame.mockRestore();
      }
      await act(async () => {
        for (const callback of frames.splice(0)) callback(performance.now());
      });
      await until(() => marker()?.dataset.version === applied);
      expect(marker()).toBe(mounted);
      return applied;
    };
    const next = await heldEdit(version, []);
    const afterEdit = layouts.slice(layouts.lastIndexOf(version) + 1);
    expect(afterEdit[0]).toBeNull();
    expect(afterEdit.at(-1)).toBe(next);
    expect(retained.toOverlayRect(unit)).toBeNull();
    const shown = geometries.at(-1)!;
    expect(shown.toOverlayRect(unit)).not.toBeNull();

    await act(async () => {
      window.dispatchEvent(new Event('resize'));
    });
    await until(() => geometries.at(-1) !== shown);
    const last = await heldEdit(next, [retained, shown]);

    // happy-dom lays out no pixels, so the pages have no client geometry to scroll to.
    expect(
      await contexts
        .at(-1)!
        .navigation.scrollToParagraph(
          { story: 'body', paraId: paragraph.paraId },
          { expectVersion: last }
        )
    ).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
  });

  test('removing the plugins releases the layout their overlays held', async () => {
    const geometries: DocxPluginGeometry[] = [];
    const plugin = defineDocxPlugin<null>({
      id: 'acme.held',
      createState: () => null,
      overlay: ({ geometry }) => {
        geometries.push(geometry);
        return null;
      },
    });
    const unit = { x: 0, y: 0, width: 1, height: 1 };
    const { rerender } = await mount({ plugins: [plugin] });
    await until(() => geometries.at(-1)?.toOverlayRect(unit) != null);
    const held = geometries.at(-1)!;
    rerender({ plugins: [] });
    await settle();
    expect(held.toOverlayRect(unit)).toBeNull();
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
    const buffer = experimentalWorkerOpen === undefined
      ? await inlineImageDocument('<w:p w14:paraId="00000001"><w:r><w:t>First paragraph</w:t></w:r></w:p>')
      : documentBytes();
    const { ref } = await mount({ plugins: [plugin] }, false, buffer);
    await until(() => geometry !== null);
    const session = experimentalWorkerOpen === false
      ? ref.current!.getEditorRef()!.getYrsSession()!
      : null;
    if (session) {
      await act(async () => {
        expect(session.persistParagraphIds().status).toBe('applied');
        ref.current!.getEditorRef()!.syncYrsInputState(true);
      });
    }
    const { version, paragraph } = await firstParagraph(ref);
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.version === version);
    const identities = session ? session.paragraphIdentities() : await ref.current!.getParagraphIdentities();
    const entry = identities
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

  test('story reads answer in a read-only viewer with one worker read each', async () => {
    const { plugin, contexts, log } = recorder();
    await mount({ plugins: [plugin], readOnly: true });
    await until(() => log.includes('load:loaded'));
    const reads: string[] = [];
    for (const worker of workers) {
      const post = worker.postMessage.bind(worker);
      worker.postMessage = (message, transfer) => {
        const read = (message as { read?: { kind?: string } }).read?.kind;
        if (read === 'listStories' || read === 'readStories') reads.push(read);
        post(message, transfer);
      };
    }
    const context = contexts.at(-1)!;
    const replicaLoads = () =>
      workers.reduce((count, worker) => count + worker.requests.filter((type) => type === 'encodeState').length, 0);
    const loadsBefore = replicaLoads();
    const listed = await context.read.listStories();
    const uses = [
      { sectionIndex: 0, variant: 'default' as const },
      { sectionIndex: 1, variant: 'default' as const },
    ];
    expect(listed.ok && listed.stories.flatMap((story) => story.uses ?? []).every(({ shown }) => shown === undefined || typeof shown === 'boolean'))
      .toBe(true);
    expect(listed.ok && listed.stories.filter((story) => story.uses).map(({ story, kind, uses }) => ({
      story, kind, uses: uses!.map(({ sectionIndex, variant }) => ({ sectionIndex, variant })),
    })))
      .toEqual([
        { story: 'hf:rId3', kind: 'header', uses },
        { story: 'hf:rId4', kind: 'footer', uses },
      ]);
    const read = await context.read.readStories({ stories: ['header', 'footer'], view: 'accepted' });
    expect(read.ok && read.stories.map((story) => [story.story, story.ok])).toEqual([
      ['hf:rId3', true],
      ['hf:rId4', true],
    ]);
    expect(reads).toEqual(experimentalWorkerOpen === false ? [] : ['listStories', 'readStories']);
    expect(replicaLoads()).toBe(loadsBefore);
  });

  test('host proposals anchor through the editor ref in a read-only viewer', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const events: DocxPluginEvent[] = [];
    const plugin = defineDocxPlugin({
      id: 'acme.proposals',
      createState: () => null,
      onEvent(_context, event) {
        if (event.type === 'proposal-change') events.push(event);
      },
      overlay: (props) => {
        geometry = props.geometry;
        return null;
      },
    });
    const { ref } = await mount({
      plugins: [plugin], readOnly: true, allowHostProposals: true,
    });
    await until(() => geometry !== null);
    const session = experimentalWorkerOpen === false
      ? ref.current!.getEditorRef()!.getYrsSession()!
      : null;
    const { version: initialVersion, paragraph } = await firstParagraph(ref);
    const identities = session ? session.paragraphIdentities() : await ref.current!.getParagraphIdentities();
    const word = paragraph.text.slice(0, 3);
    const proposed = await act(() =>
      ref.current!.proposeChanges({
        expectVersion: session ? session.version() : initialVersion,
        proposals: [
          {
            id: 'p1',
            paragraph: {
              kind: 'session',
              sessionId: identities.sessionId,
              story: paragraph.story,
              paraId: paragraph.paraId,
            },
            suggest: { author: 'Assistant', date: '2026-09-29T00:00:00Z' },
            op: 'replaceText',
            search: word,
            replaceWith: 'XYZ',
          },
        ],
      })
    );
    if (!proposed.ok) throw new Error(proposed.failure.message);
    const { version } = proposed.snapshot;
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.version === version);
    const anchorAt = (previewVersion: number) => {
      const current = geometry! as DocxPluginGeometry;
      current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
      for (const canvas of current.dom.pagesContainer.querySelectorAll('canvas[data-page-index]')) {
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
      }
      const result = current.getAnchorGeometry({ kind: 'proposal', id: 'p1' });
      expect(result).toMatchObject({ ok: true, version, previewVersion, layoutId: current.layout.id });
      if (result.ok) expect(result.rects.length).toBeGreaterThan(0);
      expect(current.getAnchorGeometry({ kind: 'proposal', id: 'missing' })).toMatchObject({
        ok: false,
        failure: { code: 'unknown-proposal' },
      });
    };
    anchorAt(0);

    const decided = await act(() =>
      ref.current!.setProposalStates({
        expectVersion: version,
        expectPreviewVersion: 0,
        changes: [{ id: 'p1', state: 'rejected' }],
      })
    );
    expect(decided).toMatchObject({ ok: true, snapshot: { version, previewVersion: 1 } });
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.previewVersion === 1);
    anchorAt(1);
    expect(events.at(-1)).toMatchObject({ type: 'proposal-change', version, previewVersion: 1 });
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

  test('a viewer opened read-only lays out inline and block reservations without loading its replica', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const buffer = await inlineImageDocument(
      '<w:p w14:paraId="00000001"><w:r><w:t>First paragraph text</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000002"><w:r><w:t>Second</w:t></w:r></w:p>'
    );
    const loadSpies: Array<ReturnType<typeof spyOn<YrsSession, 'loadState'>>> = [];
    const createSession = yrs.createYrsSession;
    const opening = spyOn(yrs, 'createYrsSession').mockImplementation(async (options) => {
      const session = await createSession(options);
      loadSpies.push(spyOn(session, 'loadState'));
      return session;
    });
    try {
      const { ref, rerender } = await mount({ readOnly: true }, false, buffer);
      const { version, paragraph } = await firstParagraph(ref);
      const replicaLoads = () => workers.reduce(
        (count, worker) => count + worker.requests.filter((type) => type === 'encodeState').length, 0
      );
      const loadsBefore = replicaLoads();
      expect(loadSpies.length).toBeGreaterThan(0);
      for (const loaded of loadSpies) loaded.mockClear();
      const plugin = defineDocxPlugin({
        id: 'host.viewer-reserve',
        createState: () => null,
        reservations: () => ({
          version,
          items: [
            {
              id: 'chip', kind: 'inline', width: 40, height: 20,
              anchor: { kind: 'range', version, range: {
                story: paragraph.story, view: 'accepted',
                start: { paraId: paragraph.paraId, offset: 6 },
                end: { paraId: paragraph.paraId, offset: 6 },
              } },
            },
            {
              id: 'gap', kind: 'block', side: 'after', height: 60,
              anchor: { kind: 'range', version, range: {
                story: paragraph.story, view: 'accepted',
                start: { paraId: paragraph.paraId, offset: 0 },
                end: { paraId: paragraph.paraId, offset: 0 },
              } },
            },
          ],
        }),
        overlay: (props) => { geometry = props.geometry; return null; },
      });
      rerender({ readOnly: true, plugins: [plugin] });
      await until(() => (geometry as DocxPluginGeometry | null)?.layout.reservationsVersion === version);
      const current = geometry! as DocxPluginGeometry;
      current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
      for (const canvas of current.dom.pagesContainer.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]')) {
        const width = Number.parseFloat(canvas.style.width);
        const height = Number.parseFloat(canvas.style.height);
        expect(width).toBeGreaterThan(0);
        expect(height).toBeGreaterThan(0);
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, width, height);
      }
      const chip = current.getAnchorGeometry({ kind: 'reservation', id: 'chip' });
      const gap = current.getAnchorGeometry({ kind: 'reservation', id: 'gap' });
      expect(chip.ok).toBe(true);
      expect(gap.ok).toBe(true);
      if (!chip.ok || !gap.ok) throw new Error('reservation geometry failed');
      expect(chip.rects).toEqual([chip.anchor]);
      expect(chip.anchor.height / chip.anchor.width).toBeCloseTo(0.5);
      expect(gap.anchor.height).toBeCloseTo(chip.anchor.height * 3);
      expect(gap.anchor.y).toBeGreaterThan(chip.anchor.y);
      expect(replicaLoads()).toBe(loadsBefore);
      for (const loaded of loadSpies) expect(loaded).not.toHaveBeenCalled();
    } finally {
      opening.mockRestore();
      for (const loaded of loadSpies) loaded.mockRestore();
    }
  });

  test('reservations take layout space geometry reports, read-only, and never reach saved bytes', async () => {
    let geometry: DocxPluginGeometry | null = null;
    const buffer = await inlineImageDocument(
      '<w:p w14:paraId="00000001"><w:r><w:t>First paragraph text</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000002"><w:r><w:t>Second</w:t></w:r></w:p>'
    );
    const { ref, rerender } = await mount({}, false, buffer);
    const { version, paragraph } = await firstParagraph(ref);
    const identity = (await ref.current!.getParagraphIdentities()).paragraphs.find(
      (candidate) => candidate.session?.paraId === paragraph.paraId
    )!.session!;
    const saved = new Uint8Array((await ref.current!.save())!);
    let evaluations = 0;
    const plugin = defineDocxPlugin({
      id: 'acme.reserve',
      createState: () => null,
      reservations: (context) => {
        evaluations += 1;
        return {
          version: context.snapshot.version,
          items: [
            {
              id: 'chip', kind: 'inline', width: 40, height: 20,
              anchor: { kind: 'search', paragraph: identity, text: 'paragraph' },
            },
            { id: 'gap', kind: 'block', side: 'after', height: 60, anchor: { kind: 'paragraph', paragraph: identity } },
          ],
        };
      },
      overlay: (props) => {
        geometry = props.geometry;
        return null;
      },
    });
    const anchored = async () => {
      await until(() => {
        const current = geometry as DocxPluginGeometry | null;
        return current?.layout.version === version && current.layout.reservationsVersion === version;
      });
      const current = geometry! as DocxPluginGeometry;
      current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
      for (const canvas of current.dom.pagesContainer.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]')) {
        const width = Number.parseFloat(canvas.style.width);
        const height = Number.parseFloat(canvas.style.height);
        expect(width).toBeGreaterThan(0);
        expect(height).toBeGreaterThan(0);
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, width, height);
      }
      const chip = current.getAnchorGeometry({ kind: 'reservation', id: 'chip' });
      const gap = current.getAnchorGeometry({ kind: 'reservation', id: 'gap' });
      if (!chip.ok || !gap.ok) throw new Error('reservation geometry failed');
      expect(chip.rects).toEqual([chip.anchor]);
      expect(chip.anchor.height / chip.anchor.width).toBeCloseTo(0.5);
      expect(gap.anchor.height).toBeCloseTo(chip.anchor.height * 3);
      expect(gap.anchor.y).toBeGreaterThan(chip.anchor.y);
      expect(current.getAnchorGeometry({ kind: 'reservation', id: 'missing' })).toMatchObject({
        ok: false, failure: { code: 'missing-target' },
      });
    };
    rerender({ plugins: [plugin] });
    await anchored();
    expect(new Uint8Array((await ref.current!.save())!)).toEqual(saved);
    const evaluated = evaluations;
    await act(async () => ref.current!.setZoom(1.25));
    await until(() => (geometry as DocxPluginGeometry | null)?.layout.zoom === 1.25);
    expect(evaluations).toBe(evaluated);
    geometry = null;
    rerender({ plugins: [plugin], readOnly: true });
    await anchored();
    rerender({ plugins: [plugin] });
    const removed = await ref.current!.applyEdits({
      expectVersion: version,
      steps: [{ op: 'deleteParagraphs', story: 'body', firstParaId: paragraph.paraId, lastParaId: paragraph.paraId }],
    });
    if (!removed.ok) throw new Error(removed.failure.message);
    await until(() => {
      const current = geometry as DocxPluginGeometry | null;
      return current?.layout.version === removed.version && current.layout.reservationsVersion === removed.version;
    });
    expect(geometry!.getAnchorGeometry({ kind: 'reservation', id: 'chip' })).toMatchObject({
      ok: false, failure: { code: 'missing-target' },
    });
  });

  test.each(['removed', 'fresh'] as const)('reservation owners update independently when another owner is stale (%s)', async (other) => {
    let access!: pluginHosts.DocxPluginHostAccess;
    const createHost = pluginHosts.createDocxPluginHost;
    const capture = spyOn(pluginHosts, 'createDocxPluginHost').mockImplementation((next) => {
      access = next;
      return createHost(next);
    });
    const buffer = await inlineImageDocument('<w:p w14:paraId="00000001"><w:r><w:t>Alpha beta gamma</w:t></w:r></w:p>');
    let mounted!: Awaited<ReturnType<typeof mount>>;
    try { mounted = await mount({}, false, buffer); } finally { capture.mockRestore(); }
    const { ref, rerender } = mounted;
    const { version, paragraph } = await firstParagraph(ref);
    const range = (atVersion: string, start: number, end = start) => ({
      kind: 'range' as const, version: atVersion, range: {
        story: 'body', view: 'accepted' as const,
        start: { paraId: paragraph.paraId, offset: start }, end: { paraId: paragraph.paraId, offset: end },
      },
    });
    const a = defineDocxPlugin({
      id: 'test.reserve-a', createState: () => null,
      reservations: () => ({ version, items: [{ id: 'a', kind: 'inline', width: 35, height: 20, anchor: range(version, 0) }] }),
    });
    const b = defineDocxPlugin({
      id: 'test.reserve-b', createState: () => null,
      reservations: (context) => ({ version: context.snapshot.version, items: [{
        id: 'b', kind: 'inline', width: context.snapshot.version === version ? 20 : 75,
        height: 20, hideAnchor: true, anchor: range(context.snapshot.version, 6, 10),
      }] }),
    });
    const geometry = () => {
      const current = access.geometry();
      if (!current) return null;
      current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
      for (const canvas of current.dom.pagesContainer.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]')) {
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, Number.parseFloat(canvas.style.width), Number.parseFloat(canvas.style.height));
      }
      return current;
    };
    const answer = (owner: string, id: string) => {
      const current = geometry();
      return current && scopePluginGeometry(current, owner).getAnchorGeometry({ kind: 'reservation', id });
    };
    rerender({ plugins: [a, b] });
    await until(() => answer(a.id, 'a')?.ok === true && answer(b.id, 'b')?.ok === true);
    expect(geometry()!.dom.getRectsForRange(7, 11).filter((rect) => rect.width > 0)).toEqual([]);
    let changed!: Awaited<ReturnType<DocxEditorRef['applyEdits']>>;
    await act(async () => { changed = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId)); });
    if (!changed.ok) throw new Error(changed.failure.message);
    rerender({ plugins: other === 'removed' ? [a] : [a, b] });
    const waitFor = async (done: () => boolean, seen: () => unknown) => {
      try {
        await until(done);
      } catch (error) {
        throw new Error(`wait failed, saw ${JSON.stringify(seen())}: ${String(error)}`);
      }
    };
    const owners = () => ({ version: geometry()?.layout.version, a: answer(a.id, 'a'), b: answer(b.id, 'b') });
    const settled = () => {
      const owner = answer(b.id, 'b');
      return other === 'removed' ? owner?.ok === false : owner?.ok === true && Math.abs(owner.anchor.width - 75) < 0.5;
    };
    await waitFor(() => geometry()?.layout.version === changed.version && answer(a.id, 'a')?.ok === true && settled(), owners);
    await act(async () => { await ref.current!.flushPendingInput(); await ref.current!.readParagraphs({ view: 'accepted' }); });
    const retained = answer(a.id, 'a');
    expect(retained?.ok).toBe(true);
    if (!retained?.ok) throw new Error('missing retained owner');
    expect(retained.anchor.width).toBeCloseTo(35);
    const updated = answer(b.id, 'b');
    if (other === 'removed') {
      expect(updated).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
      // b leaves the host before the layout without it is drawn, and both layouts share the edit's reservationsVersion.
      const drawn = () => geometry()?.dom.getRectsForRange(7, 11) ?? [];
      await waitFor(() => drawn().some((rect) => rect.width > 0), () => ({ layout: geometry()?.layout.id, widths: drawn().map((rect) => rect.width) }));
    } else {
      expect(updated?.ok).toBe(true);
      if (!updated?.ok) throw new Error('missing fresh owner');
      expect(updated.anchor.width).toBeCloseTo(75);
    }
  }, 30_000);

  test('an unchanged reservation owner stays available while another owner is removed', async () => {
    let access!: pluginHosts.DocxPluginHostAccess;
    const createHost = pluginHosts.createDocxPluginHost;
    const capture = spyOn(pluginHosts, 'createDocxPluginHost').mockImplementation((next) => {
      access = next;
      return createHost(next);
    });
    const buffer = await inlineImageDocument('<w:p w14:paraId="00000001"><w:r><w:t>Alpha beta gamma</w:t></w:r></w:p>');
    let mounted!: Awaited<ReturnType<typeof mount>>;
    try { mounted = await mount({}, false, buffer); } finally { capture.mockRestore(); }
    const { ref, rerender } = mounted;
    const { version, paragraph } = await firstParagraph(ref);
    const range = (atVersion: string, start: number, end = start) => ({
      kind: 'range' as const, version: atVersion, range: {
        story: 'body', view: 'accepted' as const,
        start: { paraId: paragraph.paraId, offset: start }, end: { paraId: paragraph.paraId, offset: end },
      },
    });
    const a = defineDocxPlugin({
      id: 'test.reserve-a', createState: () => null,
      reservations: () => ({ version, items: [{ id: 'a', kind: 'inline', width: 35, height: 20, anchor: range(version, 0) }] }),
    });
    const b = defineDocxPlugin({
      id: 'test.reserve-b', createState: () => null,
      reservations: (context) => ({ version: context.snapshot.version, items: [{
        id: 'b', kind: 'inline', width: context.snapshot.version === version ? 20 : 75,
        height: 20, hideAnchor: true, anchor: range(context.snapshot.version, 6, 10),
      }] }),
    });
    const geometry = () => {
      const current = access.geometry();
      if (!current) return null;
      current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
      for (const canvas of current.dom.pagesContainer.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]')) {
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, Number.parseFloat(canvas.style.width), Number.parseFloat(canvas.style.height));
      }
      return current;
    };
    const answer = (owner: string, id: string) => {
      const current = geometry();
      return current && scopePluginGeometry(current, owner).getAnchorGeometry({ kind: 'reservation', id });
    };
    rerender({ plugins: [a, b] });
    await until(() => answer(a.id, 'a')?.ok === true && answer(b.id, 'b')?.ok === true);
    let changed!: Awaited<ReturnType<DocxEditorRef['applyEdits']>>;
    await act(async () => { changed = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId)); });
    if (!changed.ok) throw new Error(changed.failure.message);
    await until(() => geometry()?.layout.version === changed.version && answer(a.id, 'a')?.ok === true && answer(b.id, 'b')?.ok === true);
    const laidOut = geometry()!.layout.id;
    rerender({ plugins: [a] });
    const samples: ReturnType<typeof answer>[] = [];
    for (let turn = 0; turn < 20; turn += 1) {
      samples.push(answer(a.id, 'a'));
      await Promise.resolve();
    }
    await until(() => {
      const retained = answer(a.id, 'a');
      samples.push(retained);
      const removed = answer(b.id, 'b');
      return geometry()?.layout.id !== laidOut && retained?.ok === true && removed?.ok === false && removed.failure.code === 'missing-target';
    });
    expect(samples.map((sample) => (sample?.ok ? Math.round(sample.anchor.width) : (sample?.failure.code ?? null)))).toEqual(samples.map(() => 35));
  });

  test('a removed reservation owner answers missing-target as soon as the host drops it', async () => {
    let access!: pluginHosts.DocxPluginHostAccess;
    const createHost = pluginHosts.createDocxPluginHost;
    const capture = spyOn(pluginHosts, 'createDocxPluginHost').mockImplementation((next) => {
      access = next;
      return createHost(next);
    });
    const buffer = await inlineImageDocument('<w:p w14:paraId="00000001"><w:r><w:t>Alpha beta gamma</w:t></w:r></w:p>');
    let mounted!: Awaited<ReturnType<typeof mount>>;
    try { mounted = await mount({}, false, buffer); } finally { capture.mockRestore(); }
    const { ref, rerender } = mounted;
    const { version, paragraph } = await firstParagraph(ref);
    const range = (atVersion: string, start: number, end = start) => ({
      kind: 'range' as const, version: atVersion, range: {
        story: 'body', view: 'accepted' as const,
        start: { paraId: paragraph.paraId, offset: start }, end: { paraId: paragraph.paraId, offset: end },
      },
    });
    const a = defineDocxPlugin({
      id: 'test.reserve-a', createState: () => null,
      reservations: () => ({ version, items: [{ id: 'a', kind: 'inline', width: 35, height: 20, anchor: range(version, 0) }] }),
    });
    const b = defineDocxPlugin({
      id: 'test.reserve-b', createState: () => null,
      reservations: (context) => ({ version: context.snapshot.version, items: [{
        id: 'b', kind: 'inline', width: context.snapshot.version === version ? 20 : 75,
        height: 20, hideAnchor: true, anchor: range(context.snapshot.version, 6, 10),
      }] }),
    });
    const geometry = () => {
      const current = access.geometry();
      if (!current) return null;
      current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
      for (const canvas of current.dom.pagesContainer.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]')) {
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, Number.parseFloat(canvas.style.width), Number.parseFloat(canvas.style.height));
      }
      return current;
    };
    const answer = (owner: string, id: string) => {
      const current = geometry();
      return current && scopePluginGeometry(current, owner).getAnchorGeometry({ kind: 'reservation', id });
    };
    rerender({ plugins: [a, b] });
    await until(() => answer(a.id, 'a')?.ok === true && answer(b.id, 'b')?.ok === true);
    let changed!: Awaited<ReturnType<DocxEditorRef['applyEdits']>>;
    await act(async () => { changed = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId)); });
    if (!changed.ok) throw new Error(changed.failure.message);
    await until(() => geometry()?.layout.version === changed.version && answer(a.id, 'a')?.ok === true && answer(b.id, 'b')?.ok === true);
    const laidOut = geometry()!.layout.id;
    rerender({ plugins: [a] });
    const samples: ReturnType<typeof answer>[] = [];
    for (let turn = 0; turn < 20; turn += 1) {
      samples.push(answer(b.id, 'b'));
      await Promise.resolve();
    }
    await until(() => {
      const removed = answer(b.id, 'b');
      samples.push(removed);
      return geometry()?.layout.id !== laidOut && removed?.ok === false && removed.failure.code === 'missing-target';
    });
    const codes = samples.map((sample) => (sample?.ok ? 'ok' : (sample?.failure.code ?? null)));
    expect(codes.filter((code) => code !== 'ok' && code !== 'missing-target')).toEqual([]);
  });

  test('reservation definition replacement at the same id and revision applies its new width', async () => {
    const buffer = await inlineImageDocument('<w:p w14:paraId="00000001"><w:r><w:t>Alpha beta</w:t></w:r></w:p>');
    const { ref, rerender } = await mount({}, false, buffer);
    const { version, paragraph } = await firstParagraph(ref);
    let geometry: DocxPluginGeometry | null = null;
    const definition = (width: number) => defineDocxPlugin({
      id: 'test.replace-reserve', revision: 1, createState: () => null,
      reservations: () => ({ version, items: [{ id: 'box', kind: 'inline', width, height: 20,
        anchor: { kind: 'range', version, range: { story: 'body', view: 'accepted',
          start: { paraId: paragraph.paraId, offset: 0 }, end: { paraId: paragraph.paraId, offset: 0 },
        } },
      }] }),
      overlay: (props) => { geometry = props.geometry; return null; },
    });
    const width = () => {
      const current = geometry as DocxPluginGeometry | null;
      if (!current) return null;
      current.dom.pagesContainer.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
      for (const canvas of current.dom.pagesContainer.querySelectorAll<HTMLCanvasElement>('canvas[data-page-index]')) {
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, Number.parseFloat(canvas.style.width), Number.parseFloat(canvas.style.height));
      }
      const result = current.getAnchorGeometry({ kind: 'reservation', id: 'box' });
      return result.ok ? result.anchor.width : null;
    };
    rerender({ plugins: [definition(25)] });
    await until(() => width() !== null);
    expect(width()).toBeCloseTo(25);
    rerender({ plugins: [definition(70)] });
    await until(() => width() === 70);
    expect((await firstParagraph(ref)).version).toBe(version);
  });

  test('an invalid reservation set fails its plugin with invalid-payload', async () => {
    const errors: DocxPluginError[] = [];
    const plugin = defineDocxPlugin({
      id: 'acme.bad-reserve',
      createState: () => null,
      reservations: (context) => ({
        version: context.snapshot.version,
        items: [{ id: 'gap', kind: 'hide', anchor: { kind: 'search', paragraph: null as never, text: '' } }],
      }),
    });
    await mount({ plugins: [plugin], onPluginError: (error) => errors.push(error) });
    await until(() => errors.length > 0);
    expect(errors[0]).toMatchObject({ pluginId: 'acme.bad-reserve', phase: 'reservations' });
    expect(String((errors[0]!.error as Error).message)).toBe('invalid-payload');
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
    const session = experimentalWorkerOpen === false
      ? ref.current!.getEditorRef()!.getYrsSession()!
      : null;
    const read = await ref.current!.readParagraphs({ view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    const version = session ? session.version() : read.version;
    const identities = session ? session.paragraphIdentities() : await ref.current!.getParagraphIdentities();
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
      identities
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
    let host!: pluginHosts.DocxPluginHost;
    const createHost = pluginHosts.createDocxPluginHost;
    const capture = spyOn(pluginHosts, 'createDocxPluginHost').mockImplementation((access) => {
      host = createHost(access);
      return host;
    });
    const edit = async (
      props: Partial<DocxEditorBaseProps>,
      ready: () => boolean = () => true,
      changed: () => boolean = () => true
    ) => {
      let compatibilityChanges = 0;
      const { ref, view } = await mount({
        ...props,
        ...(props.onChange ? {
          onChange: (document) => {
            compatibilityChanges += 1;
            props.onChange!(document);
          },
        } : {}),
      });
      await until(ready);
      await act(async () => { await host.notificationsIdle(); });
      const versions: string[] = [];
      const unsubscribe = ref.current!.onDocumentChange(({ version }) => versions.push(version));
      try {
        materialized.mockClear();
        compatibilityChanges = 0;
        const { version, paragraph } = await firstParagraph(ref);
        let applied!: Awaited<ReturnType<DocxEditorRef['applyEdits']>>;
        await act(async () => {
          applied = await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
        });
        expect(applied).toMatchObject({ ok: true, applied: true });
        if (!applied.ok) throw new Error(applied.failure.message);
        const committedVersion = applied.version;
        await until(() => versions.includes(committedVersion) && changed()
          && (!props.onChange || compatibilityChanges > 0));
        await act(async () => {
          await ref.current!.flushPendingInput();
          await ref.current!.readParagraphs({ view: 'accepted', paraIds: [] });
          await host.notificationsIdle();
        });
        return materialized.mock.calls.length;
      } finally {
        unsubscribe();
        view.unmount();
      }
    };
    try {
      const plain = await edit({});
      const reads: string[] = [];
      let loaded = false;
      const reader = defineDocxPlugin({
        id: 'reader',
        createState: () => null,
        async onEvent(context, event) {
          if (event.type === 'load') loaded = true;
          if (event.type !== 'document-change') return;
          const read = await context.read.readParagraphs({ view: 'accepted' });
          if (read.ok) reads.push(read.version);
        },
      });
      expect(await edit({ plugins: [reader] }, () => loaded, () => reads.length > 0)).toBe(plain);
      expect(reads).toHaveLength(1);
      let changes = 0;
      expect(await edit({ onChange: () => { changes += 1; } }, undefined, () => changes > 0))
        .toBeGreaterThan(plain);
    } finally {
      capture.mockRestore();
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

describe('DocxEditor plugins (peer-less engine)', () => {
  let canvasContext: ReturnType<typeof spyOn<HTMLCanvasElement, 'getContext'>>;
  beforeEach(() => {
    installWorkerEngine();
    canvasContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({ save() {}, restore() {}, setTransform() {},
      clearRect() {}, drawImage() {}, beginPath() {}, rect() {}, clip() {} } as unknown as CanvasRenderingContext2D);
  });
  afterEach(() => canvasContext.mockRestore());

  const mount: Mount = async (props) => {
    const ref = createRef<DocxEditorRef>();
    const view = render(<DocxEditor ref={ref} documentBuffer={documentBytes()} {...props} />);
    await until(() => ref.current?.commands.getState('save').enabled === true);
    await act(async () => {
      await ref.current!.whenLayoutComplete({ timeoutMs: 3000 });
    });
    return { ref, view };
  };

  test('the default worker emits one settled selection after keyboard navigation', async () => {
    const events: DocxPluginEvent[] = [];
    const plugin = defineDocxPlugin({ id: 'settle', createState: () => null,
      onEvent(_context, event) { events.push(event); } });
    const { view } = await mount({ plugins: [plugin] });
    await settle(400);
    events.length = 0;
    const input = view.container.querySelector<HTMLTextAreaElement>('[data-docx-worker-editor] textarea')!;
    await act(async () => {
      input.focus();
      fireEvent.keyDown(input, { key: 'ArrowRight', shiftKey: true });
    });
    await until(() => events.some((event) => event.type === 'selection-change'));
    expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(0);
    await settle(400);
    expect(events.filter((event) => event.type === 'selection-settle')).toMatchObject([{ source: 'keyboard' }]);
    expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
    await settle(400);
    expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
  });

  test('a slow worker selection mapping settles once with the pointer source and mapped range', async () => {
    const create = workerPluginGeometry.createWorkerPluginGeometry;
    let hold = false;
    let mappingHeld = false;
    let release!: () => void;
    const pending = new Promise<void>((resolve) => { release = resolve; });
    let mountedGeometry: ReturnType<typeof create> | null = null;
    const geometry = spyOn(workerPluginGeometry, 'createWorkerPluginGeometry').mockImplementation((controller, ...args) => {
      const readAt = controller.readAt;
      mountedGeometry = create({ ...controller, readAt: (async (request: Parameters<typeof readAt>[0], version: string) => {
        if (hold && request.kind === 'anchorTargets') { mappingHeld = true; await pending; }
        return readAt(request, version);
      }) as typeof readAt }, ...args);
      return mountedGeometry;
    });
    const events: DocxPluginEvent[] = [];
    const plugin = defineDocxPlugin({ id: 'slow-settle', createState: () => null,
      onEvent(_context, event) { events.push(event); } });
    try {
      const { view } = await mount({ plugins: [plugin] });
      await until(() => mountedGeometry?.selectionMapped() === true);
      await settle(400);
      events.length = 0;
      hold = true;
      const input = view.container.querySelector<HTMLTextAreaElement>('[data-docx-worker-editor] textarea')!;
      await act(async () => {
        input.focus();
        fireEvent.keyDown(input, { key: 'ArrowRight', shiftKey: true });
      });
      await until(() => mappingHeld);
      await act(async () => {
        fireEvent.pointerDown(input, { pointerType: 'mouse' });
        fireEvent.pointerUp(input, { pointerType: 'mouse' });
      });
      await settle(700);
      expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(0);
      hold = false;
      await act(async () => { release(); });
      await until(() => events.some((event) => event.type === 'selection-settle'));
      const settled = events.filter((event) => event.type === 'selection-settle');
      expect(settled).toHaveLength(1);
      expect(settled[0]!.source).toBe('pointer');
      expect(settled[0]!.selection.displayRange).not.toBeNull();
      await settle(700);
      expect(events.filter((event) => event.type === 'selection-settle')).toHaveLength(1);
    } finally { hold = false; release(); geometry.mockRestore(); }
  }, 15_000);

  test('the default worker receives the locale and one change after rerender', async () => {
    const contexts: DocxPluginContext<null>[] = [];
    const events: DocxPluginEvent[] = [];
    const plugin = defineDocxPlugin({ id: 'locale', createState: () => null,
      onEvent(context, event) { contexts.push(context); events.push(event); } });
    const buffer = documentBytes();
    const ref = createRef<DocxEditorRef>();
    const element = (locale: string) => <DocxEditor ref={ref} documentBuffer={buffer} plugins={[plugin]} i18n={{ _lang: locale }} />;
    const view = render(element('DE-de'));
    await until(() => events.some((event) => event.type === 'load'));
    expect(contexts[0]!.snapshot.locale).toBe('de-DE');
    view.rerender(element('fr'));
    await until(() => events.some((event) => event.type === 'locale-change'));
    expect(events.filter((event) => event.type === 'locale-change')).toEqual([
      { type: 'locale-change', generation: contexts.at(-1)!.snapshot.generation, locale: 'fr' },
    ]);
    expect(contexts.at(-1)!.snapshot.locale).toBe('fr');
    await act(async () => { await ref.current!.whenLayoutComplete({ timeoutMs: 3000 }); });
  });

  test('a load hook that applies a batch every run loads once and keeps its state', () =>
    loadHookAppliesOnce(mount));

  test('a change hook keeps its own batch and sets state at the applied version', async () => {
    const outcomes: unknown[] = [];
    const plugin = defineDocxPlugin<State>({
      id: 'acme.review',
      createState: () => ({ count: 0 }),
      async onEvent(context, event) {
        if (event.type !== 'document-change' || outcomes.length) return;
        const read = await context.read.readParagraphs({ view: 'accepted' });
        const paragraph = read.ok ? read.paragraphs.find((c) => c.text.length >= 3) : undefined;
        if (!read.ok || !paragraph || !context.edits) return;
        const applied = await context.edits.applyEdits(appendRequest(read.version, paragraph.paraId, '?'));
        outcomes.push(applied.ok ? true : applied.failure.code, applied.ok && context.setState({ count: 1 }, applied.version));
      },
    });
    const { ref } = await mount({ plugins: [plugin], pluginGrants: WRITE });
    const { version, paragraph } = await firstParagraph(ref);
    await act(async () => {
      await ref.current!.applyEdits(appendRequest(version, paragraph.paraId));
    });
    await until(() => outcomes.length === 2);
    await settle(150);
    expect(outcomes).toEqual([true, true]);
    expect((await firstParagraph(ref)).paragraph.text).toMatch(/[^!?]!\?$/);
  });

  test('a point read superseded by an edit returns null and the plugin keeps running', async () => {
    const errors: DocxPluginError[] = [];
    const changes: string[] = [];
    let context: DocxPluginContext<State> | undefined;
    let geometry: DocxPluginGeometry | null = null;
    const plugin = defineDocxPlugin<State>({
      id: 'point-reader',
      createState: () => ({ count: 0 }),
      initialize: value => { context = value; },
      overlay: props => { geometry = props.geometry; return null; },
      onEvent(_context, event) {
        if (event.type === 'document-change') changes.push(event.version);
      },
    });
    const { ref, view } = await mount({ plugins: [plugin], pluginGrants: { 'point-reader': { document: 'read' } }, onPluginError: error => errors.push(error) });
    const { version, paragraph } = await firstParagraph(ref);
    await until(() => geometry?.layout.version === version);
    const host = view.container.querySelector<HTMLElement>('[data-docx-pages-host]')!;
    host.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1200);
    const worker = workers.find(candidate => candidate.requests.includes('documentRead'))!;
    const reads = () => worker.requests.filter(type => type === 'documentRead').length;
    const before = reads();
    let result: unknown = 'unanswered';
    let edited = false;
    worker.hold();
    const action = context!.run(async next => {
      result = await next.geometry!.readPositionAtPoint(100, 100);
    });
    await until(() => reads() > before);
    void ref.current!.applyEdits(appendRequest(version, paragraph.paraId)).then(() => { edited = true; });
    await until(() => {
      worker.release(reply => !('read' in reply));
      return edited;
    });
    worker.release();
    await act(() => action);
    expect(result).toBeNull();
    expect(errors).toEqual([]);
    const seen = changes.length;
    const latest = await firstParagraph(ref);
    await act(async () => { await ref.current!.applyEdits(appendRequest(latest.version, latest.paragraph.paraId)); });
    await until(() => changes.length > seen);
    expect(errors).toEqual([]);
  });

  test('a read superseded by an edit refuses and the plugin keeps running', async () => {
    const errors: DocxPluginError[] = [];
    const changes: string[] = [];
    let context: DocxPluginContext<State> | undefined;
    const plugin = defineDocxPlugin<State>({
      id: 'acme.review',
      createState: () => ({ count: 0 }),
      initialize: (value) => {
        context = value;
      },
      onEvent(_context, event) {
        if (event.type === 'document-change') changes.push(event.version);
      },
    });
    const { ref } = await mount({ plugins: [plugin], pluginGrants: WRITE, onPluginError: (error) => errors.push(error) });
    await until(() => context !== undefined);
    const { version, paragraph } = await firstParagraph(ref);
    const worker = workers.find((candidate) => candidate.requests.includes('documentRead'))!;
    const reads = () => worker.requests.filter((type) => type === 'documentRead').length;
    const before = reads();
    let result: unknown;
    let edited = false;
    worker.hold();
    const action = context!.run(async (next) => {
      result = await next.read.readParagraphs({ view: 'accepted' });
    });
    await until(() => reads() > before);
    void ref.current!.applyEdits(appendRequest(version, paragraph.paraId)).then(() => {
      edited = true;
    });
    await until(() => {
      worker.release((reply) => !('read' in reply));
      return edited;
    });
    worker.release();
    await act(() => action);
    expect(result).toMatchObject({ ok: false, failure: { code: 'input-failed' } });
    expect(errors).toEqual([]);
    const seen = changes.length;
    const latest = await firstParagraph(ref);
    await act(async () => {
      await ref.current!.applyEdits(appendRequest(latest.version, latest.paragraph.paraId));
    });
    await until(() => changes.length > seen);
    expect(errors).toEqual([]);
  });
});

describe('DocxEditor plugins (in-thread engine)', () => {
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
    const identities = ref.current!.getEditorRef()!.getYrsSession()!.paragraphIdentities();
    const anchor = identities
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
});

describe('DocxEditor plugins (worker engine)', () => {
  let attachGate: Promise<void> | null = null;
  let attachRequested = false;
  const workers = setupWorkerEngine((worker) => {
    const post = worker.postMessage.bind(worker);
    worker.postMessage = (request, transfer) => {
      if (request.type === 'attachCanvases' && request.zoom === 1.5 && attachGate) {
        attachRequested = true;
        void attachGate.then(() => post(request, transfer));
        return;
      }
      post(request, transfer);
    };
  });

  async function mountWorker(props: Partial<DocxEditorBaseProps> = {}) {
    const opens = workers.reduce(
      (count, worker) => count + worker.requests.filter((type) => type === 'open').length, 0
    );
    const editor = await mount({ ...props, experimentalWorkerOpen: true });
    expect(workers.length).toBeGreaterThan(0);
    expect(workers.some(
      (worker) => worker.sessions.length > 0 && worker.requests.includes('open')
    )).toBe(true);
    expect(workers.reduce(
      (count, worker) => count + worker.requests.filter((type) => type === 'open').length, 0
    )).toBeGreaterThan(opens);
    await act(async () => {
      await editor.ref.current!.whenLayoutComplete({ timeoutMs: 3000 });
    });
    return editor;
  }

  test('overlays and layout events re-anchor once the pages show a new zoom', async () => {
    const results: DocxAnchorGeometryResult[] = [];
    const events: (DocxAnchorGeometryResult | null)[] = [];
    const availableZooms: { layout: number; dom: number }[] = [];
    let target: DocxGeometryTarget | null = null;
    const plugin = defineDocxPlugin({
      id: 'host.zoom-anchor',
      createState: () => null,
      onEvent(context, event) {
        if (event.type === 'layout-change' && event.layout?.zoom === 1.5 && target) {
          events.push(context.geometry?.getAnchorGeometry(target) ?? null);
        }
      },
      overlay: ({ geometry }) => {
        if (target) {
          const result = geometry.getAnchorGeometry(target);
          results.push(result);
          if (result.ok) availableZooms.push({ layout: geometry.layout.zoom, dom: geometry.dom.zoom });
        }
        return null;
      },
    });
    let painted = false;
    const { ref } = await mountWorker({
      plugins: [plugin],
      onFirstPagePainted: () => (painted = true),
    });
    await until(() => painted);
    const { paragraph } = await firstParagraph(ref);
    const identities = await ref.current!.getParagraphIdentities();
    const anchor = identities
      .paragraphs.find((entry) => entry.session?.paraId === paragraph.paraId)!.session!;
    target = { kind: 'paragraph', paragraph: anchor };
    const rect = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockReturnValue(
      new DOMRect(0, 0, 800, 1000)
    );
    let releaseAttach = () => {};
    attachRequested = false;
    attachGate = new Promise<void>((done) => (releaseAttach = done));
    try {
      availableZooms.length = 0;
      await act(async () => ref.current!.setZoom(1.5));
      await until(() => events.length > 0 && attachRequested);
      expect(events).toMatchObject([{ ok: false, failure: { code: 'layout-unavailable' } }]);
      expect(results.at(-1)).toMatchObject({ ok: false });
      expect(availableZooms).toEqual([]);
      await act(async () => releaseAttach());
      await until(() => {
        const last = results.at(-1);
        return !!last?.ok && last.rects.length > 0;
      });
      await until(() => events.at(-1)?.ok === true);
      expect(availableZooms.length).toBeGreaterThan(0);
      expect(availableZooms.every(({ layout, dom }) => layout === 1.5 && dom === 1.5)).toBe(true);
    } finally {
      releaseAttach();
      attachGate = null;
      rect.mockRestore();
    }
  });
});


type MatrixEngine = 'DI' | 'DL' | 'DP' | 'DV';
type MatrixSelectionPublication = { resolvedKey: string; cause?: 'remote' | 'seed' };
type MatrixSelectionDelivery = {
  event: Extract<DocxPluginEvent, { type: 'selection-change' | 'selection-settle' }>;
  publication: MatrixSelectionPublication;
  at: number;
};

function expectSelectionRange(
  selected: DocxAnchorGeometryResult,
  range: DocxAnchorGeometryResult,
  context: DocxPluginContext<null>,
  affinity: 'caret' | 'trailing' = 'caret'
) {
  if (!selected.ok) throw new Error(`Selection anchor failed: ${selected.failure.code}`);
  expect(selected.ok).toBe(true);
  expect(range.ok).toBe(true);
  if (!selected.ok || !range.ok) return;
  const { anchor: selectedAnchor, ...selectedRange } = selected;
  const { anchor: rangeAnchor, ...expectedRange } = range;
  expect(selectedRange).toEqual(expectedRange);
  if (affinity === 'trailing') {
    const last = range.rects.at(-1);
    expect(last).toBeDefined();
    const { x, ...rest } = selectedAnchor;
    expect(rest).toEqual({ y: last!.y, height: last!.height, pageIndex: last!.pageIndex, width: 0 });
    expect(x).toBeCloseTo(last!.x + last!.width, 10);
    return;
  }
  const display = context.snapshot.selection.displayRange!;
  const caret = context.geometry!.dom.getCoordinatesForPosition(display.to);
  expect(caret).not.toBeNull();
  const expectedCaret = context.geometry!.toOverlayRect({ ...caret!, width: 0 });
  expect(expectedCaret).not.toBeNull();
  expect(selectedAnchor).toEqual({ ...expectedCaret!, pageIndex: rangeAnchor.pageIndex });
}

async function matrixDocument(buffer = documentBytes()) {
  const zip = await JSZip.loadAsync(buffer);
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  let content = await zip.file('[Content_Types].xml')!.async('string');
  let relationships = await zip.file('word/_rels/document.xml.rels')!.async('string');
  let body = await zip.file('word/document.xml')!.async('string');
  for (const [kind, id] of [['footnote', 'fn'], ['endnote', 'en']] as const) {
    content = content.replace('</Types>', `<Override PartName="/word/${kind}s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/></Types>`);
    relationships = relationships.replace('</Relationships>', `<Relationship Id="matrix-${id}" Type="${rel}/${kind}s" Target="${kind}s.xml"/></Relationships>`);
    body = body.replace('</w:p>', `<w:r><w:${kind}Reference w:id="1"/></w:r></w:p>`);
    zip.file(`word/${kind}s.xml`, `<w:${kind}s xmlns:w="${w}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:${kind} w:id="1"><w:p w14:paraId="${id === 'fn' ? '60000001' : '70000001'}"><w:r><w:t>Note navigation text</w:t></w:r></w:p></w:${kind}></w:${kind}s>`);
  }
  zip.file('[Content_Types].xml', content);
  zip.file('word/_rels/document.xml.rels', relationships);
  zip.file('word/document.xml', body);
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function removeRemoteHeader(peer: YrsSession, buffer: ArrayBuffer) {
  const remote = await createResidentEngineSession(undefined, 67772);
  try {
    remote.openSourceDocx(new Uint8Array(buffer), { state: peer.encodeState() });
    expect(remote.geometryReader.hasStory('hf:rIdH1')).toBe(true);
    const document = remote.partDocument();
    const headers = new Map(document.package.headers);
    headers.delete('rIdH1');
    const relationships = new Map(document.package.relationships);
    relationships.delete('rIdH1');
    const strip = (properties: NonNullable<typeof document.package.document.finalSectionProperties>) => ({
      ...properties, headerReferences: properties.headerReferences?.filter(ref => ref.rId !== 'rIdH1'),
    });
    expect(remote.setPartDocument({ ...document, package: { ...document.package, headers, relationships,
      document: { ...document.package.document,
        sections: document.package.document.sections?.map(section => ({ ...section, properties: strip(section.properties) })),
        finalSectionProperties: strip(document.package.document.finalSectionProperties ?? {}),
      },
    } }, ['hf:rIdH1'])).toBe(true);
    peer.applyUpdate(remote.encodeStateAsUpdate(peer.encodeStateVector()));
    expect(peer.hasStory('hf:rIdH1')).toBe(false);
  } finally { remote.destroy(); }
}

async function matrixEditor(engine: MatrixEngine, buffer = documentBytes(), mapping?: { reject(): boolean; rejected(): void }, controlledClock = false) {
  const canvasContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => new Proxy({} as Record<string, unknown>, {
    get(target, key) {
      if (key === 'measureText') return () => ({ width: 0 });
      if (key === 'createImageData') return (width: number, height: number) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) });
      return key in target ? target[key as string] : () => {};
    },
    set(target, key, value) { target[key as string] = value; return true; },
  }) as unknown as CanvasRenderingContext2D);
  const bounds = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockReturnValue(new DOMRect(0, 0, 800, 1000));
  const clock = controlledClock ? selectionClock() : selectionWindow();
  const restoreClock = () => { if ('restore' in clock) clock.restore(); };
  const events: DocxPluginEvent[] = [];
  const settles: Array<{ context: DocxPluginContext<null>; anchor: DocxAnchorGeometryResult | null; at: number }> = [];
  const deliveries: MatrixSelectionDelivery[] = [];
  const publications = new WeakMap<DocxPluginSelection, MatrixSelectionPublication>();
  let publication: MatrixSelectionPublication = { resolvedKey: 'null' };
  const plugin = defineDocxPlugin({ id: 'matrix', createState: () => null,
    onEvent(context, event) {
      events.push(event);
      if (event.type === 'selection-change' || event.type === 'selection-settle') {
        const observed = event.type === 'selection-change' ? publications.get(event.selection) : publication;
        expect(observed).toBeDefined();
        deliveries.push({ event, publication: observed!, at: Date.now() });
      }
      if (event.type === 'selection-settle') settles.push({ context, at: Date.now(),
        anchor: context.geometry?.getAnchorGeometry({ kind: 'selection' }) ?? null });
    } });
  let host!: pluginHosts.DocxPluginHost;
  let controller: ReturnType<typeof shellControllers.createPeerlessShellController> | null = null;
  let workerGeometry: ReturnType<typeof workerPluginGeometry.createWorkerPluginGeometry> | null = null;
  let binding: Parameters<ReturnType<typeof shellControllers.createPeerlessShellController>['attach']>[0] | null = null;
  let replica: CollaborationReplica | null = null;
  let hookBinding: ReturnType<typeof docxBindings.useDocxPluginHost> | null = null;
  let openBand: ReturnType<typeof headerFooterEditing.useHeaderFooterEditing>['handleHeaderFooterDoubleClick'] | undefined;
  let closeBand: ReturnType<typeof headerFooterEditing.useHeaderFooterEditing>['handleBodyClick'] | undefined;
  let openNote: pagePointers.UsePagesPointerOptions['onNoteClick'];
  let restoreClearedSelection: (() => void) | undefined;
  const restores: Array<() => void> = [];
  const useBinding = docxBindings.useDocxPluginHost;
  const captureBinding = spyOn(docxBindings, 'useDocxPluginHost').mockImplementation((options) => { hookBinding = useBinding(options); return hookBinding; });
  const useParts = headerFooterEditing.useHeaderFooterEditing;
  const captureParts = spyOn(headerFooterEditing, 'useHeaderFooterEditing').mockImplementation((options) => {
    const parts = useParts(options);
    openBand = parts.handleHeaderFooterDoubleClick;
    closeBand = parts.handleBodyClick;
    return parts;
  });
  const usePointer = pagePointers.usePagesPointer;
  const capturePointer = spyOn(pagePointers, 'usePagesPointer').mockImplementation((options) => { openNote = options.onNoteClick; return usePointer(options); });
  const createHost = pluginHosts.createDocxPluginHost;
  const captureHost = spyOn(pluginHosts, 'createDocxPluginHost').mockImplementation((...args) => {
    host = createHost(...args);
    const publish = host.selectionChanged;
    host.selectionChanged = (selection, identity, source) => {
      publication = { resolvedKey: typeof identity === 'object' ? identity.resolvedKey : identity ?? JSON.stringify(selection),
        cause: source?.cause };
      publications.set(selection, publication);
      return publish(selection, identity, source);
    };
    return host;
  });
  const createGeometry = workerPluginGeometry.createWorkerPluginGeometry;
  const captureGeometry = spyOn(workerPluginGeometry, 'createWorkerPluginGeometry').mockImplementation((controller, ...args) => {
    const readAt = controller.readAt;
    const access = mapping ? { ...controller, readAt: (async (request: Parameters<typeof readAt>[0], version: string) => {
      if (mapping.reject() && request.kind === 'anchorTargets') { mapping.rejected(); throw new Error('Selection mapping rejected'); }
      return readAt(request, version);
    }) as typeof readAt } : controller;
    workerGeometry = createGeometry(access, ...args);
    return workerGeometry;
  });
  const createController = shellControllers.createPeerlessShellController;
  const captureController = spyOn(shellControllers, 'createPeerlessShellController').mockImplementation((options) => {
    controller = createController(options);
    const attach = controller.attach;
    controller.attach = (next) => { binding = next; attach(next); };
    return controller;
  });
  let mounted: { ref: React.RefObject<DocxEditorRef | null>; view: ReturnType<typeof render>; rerender(next: Partial<DocxEditorBaseProps>): void };
  let painted = false;
  const props = { plugins: [plugin], onFirstPagePainted: () => { painted = true; }, ...(engine === 'DV' ? { readOnly: true } : {}), collaboration: { onReplica: (next: CollaborationReplica | null) => { replica = next; } } };
  try {
    if (engine !== 'DI') installWorkerEngine();
    if (engine === 'DP') {
      const ref = createRef<DocxEditorRef>();
      const element = (next: Partial<DocxEditorBaseProps>) => <DocxEditor ref={ref} documentBuffer={buffer} {...props} {...next} />;
      const view = render(element({}));
      mounted = { ref, view, rerender: (next) => view.rerender(element(next)) };
      await until(() => ref.current?.commands.getState('save').enabled === true);
      await act(async () => { await ref.current!.whenLayoutComplete({ timeoutMs: 3000 }); });
    } else {
      const opened = await mountEditor(props, false, buffer, engine === 'DI' ? false : undefined);
      mounted = { ...opened, rerender: next => opened.rerender({ ...props, ...next }) };
    }
    await until(() => painted);
    await host.notificationsIdle();
    await until(() => host.activations()[0]?.context.geometry !== null && host.activations()[0]?.context.geometry !== undefined);
    await act(async () => { await clock.tick(350); });
    const context = () => host.activations()[0]!.context as DocxPluginContext<null>;
    const selection = async (story: string, paraId: string, anchor: number, head = anchor, waitForMapping = true) => {
      await act(async () => {
        restoreClearedSelection?.();
        restoreClearedSelection = undefined;
        if (controller) {
          const result = await controller.submit({ kind: 'select', payload: { selection: {
            anchor: { story, paraId, offset: anchor }, focus: { story, paraId, offset: head },
          } } });
          expect(result).toMatchObject({ ok: true });
        } else {
          const editor = mounted.ref.current!.getEditorRef()!;
          const session = editor.getYrsSession()!;
          expect(session.hasStory(story)).toBe(true);
          session.setSelection({ story, paraId, offset: anchor }, { story, paraId, offset: head });
          editor.syncYrsInputState(false);
        }
      });
      await host.notificationsIdle();
      if (workerGeometry && waitForMapping) await until(() => workerGeometry!.selectionMapped());
    };
    const quiet = async (ms = 350) => {
      await act(async () => { await clock.tick(ms); });
      await host.notificationsIdle();
    };
    const selectionLayout = async () => {
      expect(mounted.view.container.isConnected).toBe(true);
      expect(mounted.ref.current).not.toBeNull();
      await act(async () => {
        await mounted.ref.current!.flushPendingInput();
      });
      await mounted.ref.current!.whenLayoutComplete({ timeoutMs: 3000 });
      await act(async () => {});
      await new Promise<void>(resolve => {
        const ready = () => {
          const current = context();
          const display = current.snapshot.selection.displayRange;
          const layout = current.geometry?.layout;
          return !!display && !!layout && layout.version === host.version() &&
            display.layoutId === layout.id && (workerGeometry?.selectionMapped() ?? true);
        };
        const stop: Array<() => void> = [];
        const check = () => {
          if (!ready()) return;
          for (const unsubscribe of stop) unsubscribe();
          resolve();
        };
        stop.push(host.subscribe(check));
        if (workerGeometry) stop.push(workerGeometry.subscribe(check));
        check();
      });
      await host.notificationsIdle();
    };
    const reset = () => { events.length = 0; settles.length = 0; deliveries.length = 0; };
    const anchor = () => context().geometry!.getAnchorGeometry({ kind: 'selection' });
    const readAnchor = () => context().geometry!.readAnchorGeometries([{ kind: 'selection' }]);
    const input = () => {
      const surface = mounted.view.container.querySelector('[data-docx-input]');
      const id = surface?.getAttribute('data-docx-input');
      const element = id ? document.getElementById(id) : mounted.view.container.querySelector('textarea');
      expect(element).toBeInstanceOf(HTMLTextAreaElement);
      return element as HTMLTextAreaElement;
    };
    const first = (await firstParagraph(mounted.ref)).paragraph;
    reset();
    return { ...mounted, host, plugin, events, settles, deliveries, first, context, selection, selectionLayout, quiet, reset, anchor, readAnchor, clock, input,
      publication: () => publication,
      controller: () => controller,
      session: () => mounted.ref.current!.getEditorRef()!.getYrsSession()!,
      binding: () => binding,
      async remote(edit: (session: YrsSession) => void | Promise<void>, path: 'update' | 'chunks' | 'pipe' = 'update') {
        const transport: CollaborationReplica = engine === 'DI' ? this.session() : replica!;
        expect(transport).not.toBeNull();
        const peer = await createYrsSession({ clientId: 67771 });
        try {
          const access = engine === 'DP' ? (transport as AsyncReplicaAdapter).async : transport;
          peer.openDocx(new Uint8Array(buffer), false);
          peer.loadState(await access.encodeStateAsUpdate());
          const vector = await access.encodeStateVector();
          const sync = encodeSyncStep2(peer.encodeStateAsUpdate(vector)).slice().buffer;
          await edit(peer);
          const update = peer.encodeStateAsUpdate(vector);
          await act(async () => {
            if (path === 'pipe') {
              expect(binding).not.toBeNull();
              const data = encodeUpdate(update).slice().buffer;
              const ticket = binding!.session.submit('collabPipe', { provider: 84711, ops: [
                { op: 'create', seq: 1, options: {} }, { op: 'connect', seq: 2 },
                { op: 'event', seq: 3, event: { type: 'open' } },
                { op: 'event', seq: 4, event: { type: 'message', data: sync } },
                { op: 'event', seq: 5, event: { type: 'message', data } },
              ] }, { lane: 'collab', fence: true });
              expect(await ticket.ack).toMatchObject({ status: 'applied' });
            } else if (path === 'chunks') {
              expect(access.applyUpdateChunks).toBeDefined();
              await access.applyUpdateChunks!([update]);
            } else await access.applyUpdate(update);
          });
          await mounted.ref.current!.flushPendingInput();
          if (workerGeometry) await until(() => workerGeometry!.selectionMapped());
          await host.notificationsIdle();
        } finally { peer.destroy(); }
      },
      holdDisplayRange() {
        const range = spyOn(mounted.ref.current!.getEditorRef()!, 'getSelectionRange').mockReturnValue(null);
        restores.push(() => range.mockRestore());
        return { release() { range.mockRestore(); hookBinding!.publishSelection(); } };
      },
      async openStory(story: string) {
        await act(async () => {
          const kind = story.startsWith('fn:') ? 'footnote' : story.startsWith('en:') ? 'endnote' : story.includes('F') ? 'footer' : 'header';
          if (controller) expect(await controller.submit({ kind: 'openPart', payload: { kind, section: 'final', page: kind === 'header' ? 2 : 0, variant: kind === 'header' ? 'default' : 'first', noteId: 1, expectVersion: host.version() } })).toMatchObject({ ok: true });
          else if (kind === 'footnote' || kind === 'endnote') { expect(openNote).toBeDefined(); openNote!({ kind, noteId: 1 }); }
          else { expect(openBand).toBeDefined(); openBand!(kind, kind === 'header' ? 3 : 1); }
        });
      },
      async closeStory() {
        if (controller) await selection('body', first.paraId, 0);
        else await act(async () => { expect(closeBand).toBeDefined(); closeBand!(); });
      },
      async clearSelection() {
        await act(async () => {
          if (binding) {
            const view = binding.viewStore.getSnapshot();
            binding.viewStore.acceptWorkerState(view.identity, { epoch: view.identity.epoch,
              version: view.version, throughSeq: view.throughSeq, presentedThroughSeq: view.presentedThroughSeq,
              selection: null, editState: view.editState });
          } else {
            const selected = spyOn(this.session(), 'selection').mockReturnValue(null);
            const encoded = spyOn(this.session(), 'encodeSelection').mockReturnValue(null);
            const range = spyOn(mounted.ref.current!.getEditorRef()!, 'getSelectionRange').mockReturnValue(null);
            restoreClearedSelection = () => { selected.mockRestore(); encoded.mockRestore(); range.mockRestore(); };
            restores.push(() => restoreClearedSelection?.());
            mounted.ref.current!.getEditorRef()!.syncYrsInputState(false);
            hookBinding!.publishSelection();
          }
        });
        await host.notificationsIdle();
      },
      dispose() { try { mounted.view.unmount(); for (const restore of restores) restore(); canvasContext.mockRestore(); bounds.mockRestore(); } finally { restoreClock(); } },
    };
  } catch (error) { try { cleanup(); canvasContext.mockRestore(); bounds.mockRestore(); } finally { restoreClock(); } throw error; }
  finally { captureHost.mockRestore(); captureController.mockRestore(); captureBinding.mockRestore(); captureParts.mockRestore(); capturePointer.mockRestore(); captureGeometry.mockRestore(); }
}

for (const engine of ['DI', 'DL', 'DP'] as const) describe(`selection matrix ${engine}`, () => {
  for (const [location, story, paraId] of [
    ['body', 'body', '00000001'],
    ['header', 'hf:rIdH1', '10000001'],
    ['body table cell', 'body:t0:r0c0', '00000002'],
    ['header table cell', 'hf:rIdH1:t0:r0c0', '10000002'],
  ] as const) test(`a caret move in uniformly formatted ${location} text delivers selection-change to the plugin`, async () => {
    const bytes = await storyAnchorsDocx();
    const h = await matrixEditor(engine, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer, undefined, true);
    try {
      if (story.startsWith('hf:')) await h.openStory('hf:rIdH1');
      await h.selection(story, paraId, 1);
      await h.selectionLayout();
      await h.quiet();
      const previous = h.context().snapshot.selection;
      const root = story.startsWith('hf:') ? 'hf:rIdH1' : 'body';
      const input = h.controller() ? h.input()
        : h.view.container.querySelector<HTMLTextAreaElement>(`textarea[data-yrs-story="${root}"]`);
      expect(input).toBeInstanceOf(HTMLTextAreaElement);
      h.reset();
      await act(async () => {
        input!.focus();
        fireEvent.keyDown(input!, { key: 'ArrowRight' });
        await h.ref.current!.flushPendingInput();
      });
      if (h.controller()) expect(h.controller()!.getSnapshot().view!.workerSelection!.focus).toMatchObject({ story, paraId, offset: 2 });
      else expect(h.session().selection()!.head).toMatchObject({ story, paraId, offset: 2 });
      await h.selectionLayout();
      await h.host.notificationsIdle();
      const changes = h.deliveries.filter(delivery => delivery.event.type === 'selection-change');
      expect(changes.length).toBeGreaterThan(0);
      expect(changes.at(-1)!.event.selection).toMatchObject({
        formatting: previous.formatting,
        displayRange: { story },
      });
      const range = changes.at(-1)!.event.selection.displayRange!;
      expect(range.from).toBe(range.to);
      expect(range.from).not.toBe(previous.displayRange!.from);
      expect(changes.at(-1)!.event.selection).toEqual(h.context().snapshot.selection);
      expect(h.events.filter(event => event.type === 'selection-settle')).toEqual([]);
    } finally { h.dispose(); }
  }, 15_000);

  test('H2: every changed local publication settles exactly once across mixed stories and remote edits', async () => {
    const bytes = await storyAnchorsDocx();
    const source = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const h = await matrixEditor(engine, await matrixDocument(source), undefined, true);
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      let previous = h.publication().resolvedKey;
      h.reset();
      const actions: Array<() => Promise<void>> = [
        () => h.selection('body', h.first.paraId, 2, 4),
        () => h.selection('body', h.first.paraId, 2, 4),
        () => h.remote(peer => peer.insertText({ story: 'body', paraId: h.first.paraId, offset: 0 }, 'remote ')),
        () => h.selection('body', h.first.paraId, 2, 4),
        () => h.openStory('hf:rIdH1'),
        () => h.selection('hf:rIdH1', '10000001', 1, 3),
        () => h.selection('hf:rIdH1:t0:r0c0', '10000002', 1, 4),
        () => h.openStory('hf:rIdF1'),
        () => h.selection('hf:rIdF1', '10000021', 1, 3),
        () => h.clearSelection(),
        () => h.selection('hf:rIdF1', '10000021', 1, 3),
        () => h.openStory('fn:1'),
        () => h.selection('fn:1', '60000001', 1, 3),
        () => h.selection('fn:1', '60000001', 4, 6),
        () => h.openStory('en:1'),
        () => h.selection('en:1', '70000001', 1, 3),
        () => h.closeStory(),
        () => h.selection('body:t0:r0c0', '00000002', 1, 4),
        () => h.remote(peer => peer.insertText({ story: 'body', paraId: h.first.paraId, offset: 0 }, 'after '), engine === 'DP' ? 'pipe' : 'chunks'),
        () => h.selection('body', h.first.paraId, 1, 3),
        () => h.clearSelection(),
        () => h.selection('body', h.first.paraId, 1, 3),
      ];
      let localChanges = 0;
      let remoteChanges = 0;
      for (const action of actions) {
        expect(h.view.container.isConnected).toBe(true);
        expect(h.ref.current).not.toBeNull();
        const start = h.deliveries.length;
        await action();
        const [positions, exists] = JSON.parse(h.publication().resolvedKey) as [unknown[] | null, boolean];
        const firstPosition = positions?.[0];
        const story = Array.isArray(firstPosition) ? firstPosition[0] : firstPosition;
        if (exists && typeof story === 'string' && (story.startsWith('body') || story.startsWith('hf:'))) {
          await h.selectionLayout();
        }
        await h.host.notificationsIdle();
        await h.quiet(349);
        expect(h.deliveries.slice(start).filter(delivery => delivery.event.type === 'selection-settle')).toEqual([]);
        await h.quiet(1);
        const delivered = h.deliveries.slice(start);
        const expected: MatrixSelectionDelivery[] = [];
        for (const delivery of delivered) {
          if (delivery.event.type !== 'selection-change') continue;
          if (delivery.publication.resolvedKey !== previous) {
            if (delivery.publication.cause === 'remote') remoteChanges += 1;
            else if (delivery.publication.cause !== 'seed') {
              expected.push(delivery);
              localChanges += 1;
            }
          }
          previous = delivery.publication.resolvedKey;
        }
        const settled = delivered.filter(delivery => delivery.event.type === 'selection-settle');
        expect(settled.map(delivery => delivery.publication.resolvedKey))
          .toEqual(expected.map(delivery => delivery.publication.resolvedKey));
        for (const [index, delivery] of settled.entries()) {
          const change = expected[index]!;
          expect(delivery.at - change.at).toBeGreaterThanOrEqual(350);
          const latest = delivered.filter(candidate => candidate.event.type === 'selection-change' &&
            candidate.publication.resolvedKey === delivery.publication.resolvedKey).at(-1)!;
          expect(delivery.event.selection).toEqual(latest.event.selection);
        }
        await h.quiet(700);
        expect(h.deliveries.slice(start).filter(delivery => delivery.event.type === 'selection-settle')).toHaveLength(expected.length);
      }
      expect(localChanges).toBeGreaterThanOrEqual(16);
      expect(remoteChanges).toBeGreaterThanOrEqual(1);
    } finally { h.dispose(); }
  }, 30_000);

  test('a no-op local undo keeps the selection settled without a second settle', async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      expect(h.context().snapshot.selectionSettled).toBe(true);
      const version = h.context().snapshot.version;
      const range = h.context().snapshot.selection.displayRange;
      h.reset();
      await act(async () => {
        if (h.controller()) {
          expect(h.controller()!.getSnapshot().view!.editState!.canUndo).toBe(false);
          expect(await h.controller()!.submit({ kind: 'undo', payload: {} })).toMatchObject({ ok: true });
        } else {
          const editor = h.ref.current!.getEditorRef()!;
          expect(editor.canUndo()).toBe(false);
          expect(editor.undo()).toBe(false);
        }
      });
      expect(h.context().snapshot.selectionSettled).toBe(true);
      await h.quiet(349);
      expect(h.settles).toHaveLength(0);
      expect(h.context().snapshot.selectionSettled).toBe(true);
      await h.quiet(1);
      expect(h.settles).toHaveLength(0);
      expect(h.context().snapshot.selectionSettled).toBe(true);
      expect(h.context().snapshot.version).toBe(version);
      expect(h.context().snapshot.selection.displayRange).toMatchObject({ story: range!.story, from: range!.from, to: range!.to });
      await h.quiet(700);
      expect(h.settles).toHaveLength(0);
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') test('a layout refresh leaves the selection settled', async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      expect(h.context().snapshot.selectionSettled).toBe(true);
      const layout = h.context().snapshot.layout!.id;
      const states: boolean[] = [];
      const stop = h.host.subscribe(() => states.push(h.context().snapshot.selectionSettled));
      h.reset();
      try {
        await act(async () => {
          h.ref.current!.setZoom(1.25);
          if (h.controller()) expect(await h.controller()!.submit({ kind: 'setSectionProperties', payload: {
            section: 'final', patch: { pageWidth: 10000 }, expectVersion: h.host.version(),
          } })).toMatchObject({ ok: true });
          else {
            h.session().setSectionProperties(h.session().listSections().length, { pageWidth: 10000 });
            h.ref.current!.getEditorRef()!.syncYrsInputState(true);
          }
        });
        await h.selectionLayout();
        await until(() => h.context().snapshot.layout?.zoom === 1.25);
        expect(h.context().snapshot.layout!.id).not.toBe(layout);
        await h.quiet(700);
        expect(h.events.some(event => event.type === 'layout-change')).toBe(true);
        expect(states.length).toBeGreaterThan(0);
        expect(states).not.toContain(false);
        expect(h.settles).toHaveLength(0);
        expect(h.context().snapshot.selectionSettled).toBe(true);
        expect((await h.readAnchor())[0]).toMatchObject({ ok: true });
      } finally { stop(); }
    } finally { h.dispose(); }
  });

  if (engine === 'DP') test('T1: a rejected selection mapping ends waiting and settles once', async () => {
    let rejectMapping = false;
    let rejected = false;
    const h = await matrixEditor(engine, documentBytes(), { reject: () => rejectMapping, rejected: () => { rejected = true; } });
    try {
      h.reset();
      rejectMapping = true;
      await h.selection('body', h.first.paraId, 1, 3, false);
      await until(() => rejected);
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      expect(h.context().snapshot.selectionSettled).toBe(true);
      expect(h.settles[0]!.anchor).toMatchObject({ ok: false, failure: { code: 'read-failed' } });
      expect(await h.readAnchor()).toMatchObject([{ ok: false, failure: { code: 'read-failed' } }]);
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
    } finally { h.dispose(); }
  });

  if (engine !== 'DP') for (const outcome of ['mapped', 'expired'] as const) test(`F: not-ready selection is pending and its held display read ${outcome}`, async () => {
    const h = await matrixEditor(engine);
    try {
      const held = h.holdDisplayRange();
      await h.selection('body', h.first.paraId, 1, 3);
      expect(h.anchor()).toMatchObject({ ok: false, failure: { code: 'pending' } });
      let answer: readonly DocxAnchorGeometryResult[] | undefined;
      const read = h.readAnchor().then(value => { answer = value; });
      await act(async () => {});
      expect(answer).toBeUndefined();
      if (outcome === 'mapped') await act(async () => { held.release(); });
      else await act(async () => { await h.clock.expire(() => answer !== undefined); });
      await h.host.notificationsIdle();
      expect(answer).toBeDefined();
      expect(answer![0]).toMatchObject(outcome === 'mapped' ? { ok: true } : { ok: false, failure: { code: 'layout-unavailable' } });
      await read;
      if (outcome === 'expired') await act(async () => { held.release(); });
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') test('T2: a direction-only reversal settles once and refuses retained geometry', async () => {
    const h = await matrixEditor(engine);
    try {
      await h.selection('body', h.first.paraId, 0, 3);
      await h.quiet();
      const retained = h.context();
      expect(h.anchor()).toMatchObject({ ok: true });
      h.reset();
      await h.selection('body', h.first.paraId, 3, 0);
      expect(h.context().snapshot.selectionSettled).toBe(false);
      expect(retained.geometry!.getAnchorGeometry({ kind: 'selection' }))
        .toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
      expect(await retained.geometry!.readAnchorGeometries([{ kind: 'selection' }]))
        .toMatchObject([{ ok: false, failure: { code: 'layout-unavailable' } }]);
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      const result = h.anchor();
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.anchor.x).toBeCloseTo(result.rects[0]!.x, 3);
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
    } finally { h.dispose(); }
  });

  const remotePaths: readonly ('update' | 'chunks' | 'pipe')[] = engine === 'DP' ? ['update', 'chunks', 'pipe'] : ['update', 'chunks'];
  for (const path of remotePaths) for (const place of ['elsewhere', 'after', 'before', 'inside', 'collapse', 'delete'] as const) test(`${path}: ${place === 'elsewhere' ? 'T4' : place === 'after' ? 'T4a' : place === 'before' ? 'T4b' : 'T4c'}: a remote edit ${place === 'elsewhere' ? 'elsewhere' : place === 'delete' ? 'that deletes the selection' : `${place} the selection`} ${place === 'delete' || place === 'collapse' ? 'settles once' : 'neither unsettles nor settles'}`, async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    const destroyed = place === 'delete' || place === 'collapse';
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      const beforeRange = h.context().snapshot.selection.displayRange;
      const states: boolean[] = [];
      const stop = h.host.subscribe(() => states.push(h.context().snapshot.selectionSettled));
      h.reset();
      try {
        await h.remote(peer => {
          if (place === 'elsewhere') {
            const paragraph = peer.paragraphs('body').find(candidate => candidate.paraId !== h.first.paraId)!;
            expect(paragraph).toBeDefined();
            peer.insertText({ story: 'body', paraId: paragraph.paraId, offset: paragraph.text.length }, 'remote');
          } else if (place === 'collapse') peer.deleteRange({ story: 'body', start: { paraId: h.first.paraId, offset: 1 }, end: { paraId: h.first.paraId, offset: 3 } });
          else if (place === 'delete') peer.deleteRange({ story: 'body', start: { paraId: h.first.paraId, offset: 0 }, end: { paraId: h.first.paraId, offset: h.first.text.length } });
          else peer.insertText({ story: 'body', paraId: h.first.paraId, offset: place === 'before' ? 0 : place === 'inside' ? 2 : h.first.text.length }, 'remote ');
        }, path);
        await h.selectionLayout();
        await h.quiet();
        expect(h.settles).toHaveLength(destroyed ? 1 : 0);
        if (!destroyed) expect(states).not.toContain(false);
        expect(h.context().snapshot.selectionSettled).toBe(true);
        if (!destroyed) {
          expect(h.events.some(event => event.type === 'document-change')).toBe(true);
          await act(async () => { await h.ref.current!.whenLayoutComplete({ timeoutMs: 3000 }); });
          expect((await h.readAnchor())[0]).toMatchObject({ ok: true });
          expect((await h.ref.current!.readSelectionInfo())?.selectedText).toBe(place === 'inside'
            ? h.first.text.slice(1, 2) + 'remote ' + h.first.text.slice(2, 3) : h.first.text.slice(1, 3));
          expect(h.events.some(event => event.type === 'selection-change' || event.type === 'layout-change')).toBe(true);
          if (place === 'before') expect(h.context().snapshot.selection.displayRange!.from).toBeGreaterThan(beforeRange!.from);
        } else {
          await h.quiet(700);
          expect(h.settles).toHaveLength(1);
        }
      } finally { stop(); }
    } finally { h.dispose(); }
  });

  for (const path of remotePaths) test(`${path}: removing the selected story settles once`, async () => {
    const bytes = await storyAnchorsDocx();
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const h = await matrixEditor(engine, buffer, undefined, true);
    try {
      await h.openStory('hf:rIdH1');
      await h.selection('hf:rIdH1', '10000001', 1, 3);
      await h.selectionLayout();
      await h.quiet();
      h.reset();
      await h.remote(async peer => {
        expect(peer.hasStory('hf:rIdH1')).toBe(true);
        expect(peer.paragraphs('hf:rIdH1').map(paragraph => paragraph.paraId)).toContain('10000001');
        await removeRemoteHeader(peer, buffer);
      }, path);
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      expect(h.context().snapshot.selectionSettled).toBe(true);
      expect(h.context().snapshot.selection.displayRange?.story).not.toBe('hf:rIdH1');
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
    } finally { h.dispose(); }
  });

  for (const path of remotePaths) test(`${path}: a local return after a silent remote move settles again`, async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      const a = h.context().snapshot.selection.displayRange!;
      h.reset();
      await h.remote(peer => peer.insertText({ story: 'body', paraId: h.first.paraId, offset: 0 }, 'remote '), path);
      await h.selectionLayout();
      await h.quiet();
      expect(h.settles).toHaveLength(0);
      expect(h.context().snapshot.selection.displayRange!.from).toBeGreaterThan(a.from);
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      expect(h.settles[0]!.context.snapshot.selection.displayRange).toMatchObject({ story: a.story, from: a.from, to: a.to });
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
    } finally { h.dispose(); }
  });

  test('reader and seeding publications keep the selection settled', async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      h.reset();
      const states: boolean[] = [];
      const stop = h.host.subscribe(() => states.push(h.context().snapshot.selectionSettled));
      try {
        await act(async () => {
          await h.ref.current!.readSelectionInfo();
          await h.ref.current!.readParagraphs({ view: 'accepted', paraIds: [h.first.paraId] });
          if (!h.controller()) {
            const current = h.session().selection()!;
            withYrsSelectionCause(h.session(), 'seed', () => h.session().setSelection(current.anchor, current.head));
            h.ref.current!.getEditorRef()!.getSelectionRange();
          }
        });
        await h.quiet(700);
        expect(h.settles).toHaveLength(0);
        expect(states).not.toContain(false);
        expect(h.context().snapshot.selectionSettled).toBe(true);
      } finally { stop(); }
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') for (const family of ['typing', 'keyboard navigation', 'pointer', 'host setSelection', 'table insert', 'image insert', 'hyperlink insert', 'undo', 'redo', 'story switch'] as const) test(`H2: ${family} settles once without a path revision bump`, async () => {
    const bytes = family === 'story switch' ? await storyAnchorsDocx() : null;
    const buffer = bytes ? bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer : documentBytes();
    const h = await matrixEditor(engine, buffer, undefined, true);
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.selectionLayout();
      await h.quiet();
      const insertText = async () => {
        if (h.controller()) expect(await h.controller()!.submit({ kind: 'insertText', payload: { text: 'x' } })).toMatchObject({ ok: true });
        else h.ref.current!.getEditorRef()!.insertText('x');
      };
      if (family === 'undo' || family === 'redo') {
        await act(insertText);
        await h.selectionLayout();
        await h.quiet();
        if (family === 'redo') {
          await act(async () => { expect(await h.ref.current!.commands.execute('undo', null)).toMatchObject({ ok: true }); });
          await h.selectionLayout();
          await h.quiet();
        }
      }
      h.reset();
      const before = h.context().snapshot.selection.displayRange!;
      if (family === 'story switch') await h.openStory('hf:rIdH1');
      else await act(async () => {
        const controller = h.controller();
        const editor = controller ? null : h.ref.current!.getEditorRef();
        if (family === 'typing') {
          fireEvent.keyDown(h.input(), { key: 'x' });
          await insertText();
        } else if (family === 'keyboard navigation') {
          fireEvent.keyDown(h.input(), { key: 'ArrowRight' });
        } else if (family === 'pointer') {
          const target = h.view.container.querySelector<HTMLElement>('[data-page-index="0"] canvas, canvas[data-page-index="0"]')!;
          expect(target).not.toBeNull();
          const root = h.view.container.querySelector<HTMLElement>(controller ? '[data-docx-worker-editor]' : '.paged-editor');
          expect(root).not.toBeNull();
          const pointerTarget = controller ? target : root!;
          expect(root!.contains(pointerTarget)).toBe(true);
          const pages = target.closest<HTMLElement>('[data-docx-pages-host]');
          if (pages) Object.defineProperties(pages, {
            setPointerCapture: { configurable: true, value() {} },
            hasPointerCapture: { configurable: true, value: () => true },
            releasePointerCapture: { configurable: true, value() {} },
          });
          const anchor = h.anchor();
          expect(anchor.ok).toBe(true);
          if (!anchor.ok) throw new Error('Pointer target is unavailable');
          const point = { pointerId: 1, pointerType: 'mouse', button: 0, buttons: 1,
            clientX: anchor.anchor.x, clientY: anchor.anchor.y + anchor.anchor.height / 2, detail: 1 };
          fireEvent.pointerDown(pointerTarget, point);
          if (!controller) fireEvent.mouseDown(target, point);
          fireEvent.mouseUp(document, { ...point, buttons: 0 });
          fireEvent.pointerUp(pointerTarget, { ...point, buttons: 0 });
        } else if (family === 'host setSelection') {
          if (controller) expect(await controller.submit({ kind: 'select', payload: { selection: {
            anchor: { story: 'body', paraId: h.first.paraId, offset: 4 },
            focus: { story: 'body', paraId: h.first.paraId, offset: 4 },
          } } })).toMatchObject({ ok: true });
          else editor!.setSelection(before.from + 3, before.from + 3);
        } else if (family === 'table insert') {
          if (controller) expect(await controller.submit({ kind: 'insertTable', payload: { rows: 2, columns: 2 } })).toMatchObject({ ok: true });
          else expect(editor!.applyYrsCommand({ type: 'insertTable', rows: 2, columns: 2 })).toBe(true);
        } else if (family === 'image insert') {
          const image = { rId: 'synthetic-image', widthEmu: 952500, heightEmu: 952500 };
          if (controller) expect(await controller.submit({ kind: 'insertImage', payload: { image } })).toMatchObject({ ok: true });
          else expect(editor!.applyYrsCommand({ type: 'insertImage', image })).toBe(true);
        } else if (family === 'hyperlink insert') {
          if (controller) expect(await controller.submit({ kind: 'setHyperlink', payload: {
            selection: controller.getSnapshot().view!.workerSelection!, hyperlink: { href: 'https://example.com' },
            displayText: 'link', expectVersion: h.host.version(),
          } })).toMatchObject({ ok: true });
          else expect(editor!.applyYrsCommand({ type: 'setHyperlink', href: 'https://example.com', displayText: 'link' })).toBe(true);
        } else expect(await h.ref.current!.commands.execute(family, null)).toMatchObject({ ok: true });
        await h.ref.current!.flushPendingInput();
      });
      await h.selectionLayout();
      expect(h.context().snapshot.selection.displayRange).not.toMatchObject({ story: before.story, from: before.from, to: before.to });
      expect(h.context().snapshot.selectionSettled).toBe(false);
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      if (family === 'pointer') expect(h.events.find(event => event.type === 'selection-settle')).toMatchObject({ source: 'pointer' });
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
    } finally { h.dispose(); }
  });

  if (engine === 'DP') test('H2: mapped selection waits while a local intent is unacknowledged', async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    const releases: Array<() => void> = [];
    const intercepts = workers.map(worker => {
      const post = worker.postMessage.bind(worker);
      return spyOn(worker, 'postMessage').mockImplementation((request, transfer) => {
        if (request.type === 'edit' && request.command.kind === 'insertText') releases.push(() => post(request, transfer));
        else post(request, transfer);
      });
    });
    let pending: ReturnType<NonNullable<ReturnType<typeof h.controller>>['submit']> | undefined;
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      h.reset();
      await act(async () => { pending = h.controller()!.submit({ kind: 'insertText', payload: { text: 'held' } }); });
      await until(() => releases.length > 0);
      expect(h.binding()!.admission.getSnapshot().originalUnackedIntents).toBeGreaterThan(0);
      await h.quiet(700);
      expect(h.settles).toHaveLength(0);
      expect(h.context().snapshot.selectionSettled).toBe(false);
      await act(async () => { for (const release of releases.splice(0)) release(); await pending; });
      await h.selectionLayout();
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      expect(h.context().snapshot.selectionSettled).toBe(true);
    } finally {
      for (const intercept of intercepts) intercept.mockRestore();
      for (const release of releases.splice(0)) release();
      await pending;
      h.dispose();
    }
  });

  if (engine !== 'DL') test('T5: undo after replacing a selection settles once with a collapsed selection', async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    try {
      await h.selection('body', h.first.paraId, 0, 3);
      await h.selectionLayout();
      await h.quiet();
      await act(async () => {
        if (h.controller()) expect(await h.controller()!.submit({ kind: 'replaceSelection', payload: { text: 'replacement' } })).toMatchObject({ ok: true });
        else h.ref.current!.getEditorRef()!.insertText('replacement');
      });
      await h.selectionLayout();
      await h.quiet();
      h.reset();
      await act(async () => { expect(await h.ref.current!.commands.execute('undo', null)).toMatchObject({ ok: true }); });
      await h.selectionLayout();
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      expect(h.context().snapshot.selectionSettled).toBe(true);
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
      expect((await h.ref.current!.readSelectionInfo())?.selectedText ?? '').toBe('');
    } finally { h.dispose(); }
  });

  if (engine === 'DI') test('T6: focus change neither cancels nor restarts a pending settle', async () => {
    const h = await matrixEditor(engine);
    try {
      h.reset();
      const selectedAt = Date.now();
      await h.selection('body', h.first.paraId, 0, 3);
      await h.quiet(200);
      const input = h.input();
      const focusedAt = Date.now();
      await act(async () => { input.focus(); input.blur(); });
      expect(h.settles).toHaveLength(0);
      await until(() => h.settles.length > 0);
      expect(h.settles).toHaveLength(1);
      expect(h.settles[0]!.at - selectedAt).toBeGreaterThanOrEqual(350);
      expect(h.settles[0]!.at - focusedAt).toBeLessThan(350);
      expect(h.events.some(event => event.type === 'focus-change')).toBe(true);
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') test('T7: reload drops a pending settle and refuses retained geometry', async () => {
    const h = await matrixEditor(engine);
    try {
      await h.selection('body', h.first.paraId, 0, 3);
      const retained = h.context();
      const generation = retained.snapshot.generation;
      h.reset();
      let loading!: Promise<void>;
      let loaded = false;
      let loadError: unknown;
      await act(async () => {
        loading = h.ref.current!.loadDocumentBuffer(documentBytes());
        void loading.then(() => { loaded = true; }, error => { loadError = error; });
      });
      await until(() => loaded || loadError !== undefined);
      await loading;
      await h.quiet(700);
      expect(h.events.filter(event => event.type === 'selection-settle' && event.generation === generation)).toEqual([]);
      expect(h.host.generation()).not.toBe(generation);
      expect(retained.geometry!.getAnchorGeometry({ kind: 'selection' })).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
    } finally { h.dispose(); }
  });

  if (engine === 'DI') test('T8: removing one of two plugins keeps the other pending settle; reattach does not replay', async () => {
    const h = await matrixEditor(engine);
    const otherEvents: DocxPluginEvent[] = [];
    const other = defineDocxPlugin({ id: 'other', createState: () => null, onEvent(_context, event) { otherEvents.push(event); } });
    try {
      h.rerender({ plugins: [h.plugin, other] });
      await h.host.notificationsIdle();
      h.reset();
      await h.selection('body', h.first.paraId, 0, 3);
      await h.quiet(200);
      h.rerender({ plugins: [other] });
      await h.quiet(150);
      expect(h.settles).toHaveLength(0);
      expect(otherEvents.filter(event => event.type === 'selection-settle')).toHaveLength(1);
      h.rerender({ plugins: [h.plugin, other] });
      await h.quiet(700);
      expect(h.settles).toHaveLength(0);
      expect(otherEvents.filter(event => event.type === 'selection-settle')).toHaveLength(1);
    } finally { h.dispose(); }
  });

  if (engine === 'DL') test('T9: legacy body range and caret settle with their anchor', async () => {
    const h = await matrixEditor(engine);
    try {
      for (const head of [3, 0]) {
        h.reset();
        await h.selection('body', h.first.paraId, 0, head);
        await h.quiet();
        expect(h.settles).toHaveLength(1);
        expect(h.settles[0]!.anchor).toMatchObject({ ok: true, anchor: { width: 0 } });
        if (head === 0) expect(h.settles[0]!.anchor).toMatchObject({ rects: [] });
        expect((await h.readAnchor())[0]).toEqual(h.anchor());
      }
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') for (const story of ['fn:1', 'en:1']) {
    test(`T3: moving within one note settles again (${story})`, async () => {
      const h = await matrixEditor(engine, await matrixDocument());
      try {
        const paraId = story === 'fn:1' ? '60000001' : '70000001';
        await h.openStory(story);
        await h.selection(story, paraId, 1, 3);
        await h.quiet();
        h.reset();
        await h.selection(story, paraId, 4, 6);
        expect(h.context().snapshot.selectionSettled).toBe(false);
        await h.quiet();
        expect(h.settles).toHaveLength(1);
        expect(h.anchor()).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
      } finally { h.dispose(); }
    });
    test(`T13: real note navigation settles with unsupported anchor (${story})`, async () => {
      const h = await matrixEditor(engine, await matrixDocument());
      try {
        h.reset();
        await h.openStory(story);
        await h.selection(story, story === 'fn:1' ? '60000001' : '70000001', 1, 3);
        await h.quiet();
        expect(h.settles).toHaveLength(1);
        expect(h.anchor()).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
        expect(await h.readAnchor()).toMatchObject([{ ok: false, failure: { code: 'unsupported' } }]);
      } finally { h.dispose(); }
    });
  }

  if (engine !== 'DL') test('T14: selection in a body and a header table cell anchors through its root', async () => {
    const bytes = await storyAnchorsDocx();
    const h = await matrixEditor(engine, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    try {
      for (const [story, paraId] of [['body:t0:r0c0', '00000002'], ['hf:rIdH1:t0:r0c0', '10000002']]) {
        h.reset();
        if (story!.startsWith('hf:')) await h.openStory(story!);
        await h.selection(story!, paraId!, 1, 4);
        await h.quiet();
        expect(h.settles).toHaveLength(1);
        const selected = (await h.readAnchor())[0]!;
        expect(selected).toMatchObject({ ok: true });
        const [range] = await h.context().geometry!.readAnchorGeometries([{ kind: 'range', version: h.context().snapshot.version,
          range: { story: story!, start: { paraId: paraId!, offset: 1 }, end: { paraId: paraId!, offset: 4 }, view: 'accepted' } }]);
        expect(selected).toEqual(range!);
      }
    } finally { h.dispose(); }
  });

  if (engine !== 'DI') for (const [story, paraId] of [['hf:rIdH1', '10000001'], ['hf:rIdF1', '10000021']] as const) {
    test(`T11: ${story} settles once with its focus anchor`, async () => {
      const bytes = await storyAnchorsDocx();
      const h = await matrixEditor(engine, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
      try {
        await h.openStory(story);
        await h.quiet();
        h.reset();
        await h.selection(story, paraId, 1, 4);
        await h.quiet();
        expect(h.settles).toHaveLength(1);
        expect(h.settles[0]!.context.snapshot.selection.displayRange?.story).toBe(story);
        expect((await h.readAnchor())[0]).toMatchObject({ ok: true, anchor: { width: 0 } });
      } finally { h.dispose(); }
    });
  }

  if (engine !== 'DL') test('T12: story switch settles once with the new story anchor', async () => {
    const bytes = await storyAnchorsDocx();
    const h = await matrixEditor(engine, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    try {
      await h.selection('body', h.first.paraId, 1, 3);
      await h.quiet();
      const retained = h.context();
      h.reset();
      await h.openStory('hf:rIdH1');
      await h.selection('hf:rIdH1', '10000001', 1, 3);
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      expect(h.settles[0]!.context.snapshot.selection.displayRange?.story).toBe('hf:rIdH1');
      expect((await h.readAnchor())[0]).toMatchObject({ ok: true, anchor: { width: 0 } });
      expect(retained.geometry!.getAnchorGeometry({ kind: 'selection' }))
        .toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') test.each(['image', 'shape'] as const)('T16/B11: a %s selection exposes its text atom range', async (kind) => {
    const shape = '<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="shape" style="width:100pt;height:40pt"><v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox></v:shape></w:pict></w:r><w:r><w:t> after</w:t></w:r></w:p>';
    const buffer = await inlineImageDocument(kind === 'image' ? undefined : shape);
    const originalImage = globalThis.Image;
    const h = await (async () => {
      try {
        if (engine === 'DI' && kind === 'image') {
          class LoadedImage {
            naturalWidth = 1;
            naturalHeight = 1;
            onload: (() => void) | null = null;
            onerror: (() => void) | null = null;
            set src(_value: string) { queueMicrotask(() => this.onload?.()); }
          }
          globalThis.Image = LoadedImage as unknown as typeof Image;
        }
        return await matrixEditor(engine, buffer);
      } finally { globalThis.Image = originalImage; }
    })();
    try {
      expect(h.first.paraId).toBe('00000001');
      expect(h.first.text.slice(7, 8)).toBe('\uFFFC');
      if (kind === 'image' && h.controller()) {
        const objects = h.controller()!.getSnapshot().view!.geometry!.objects!;
        const picture = objects.objects.find(object => object.kind === 'image')!;
        expect(picture).toBeDefined();
        await act(async () => { expect(await h.controller()!.submit({ kind: 'select', payload: { object: {
          handle: picture.handle, expectVersion: objects.version,
        } } })).toMatchObject({ ok: true }); });
        expect(h.controller()!.getSnapshot().view!.workerSelection).toMatchObject({
          anchor: { story: 'body', paraId: '00000001', offset: 7 }, focus: { story: 'body', paraId: '00000001', offset: 8 },
        });
      }
      await h.selection('body', '00000001', 7, 8);
      await h.selectionLayout();
      await h.quiet();
      if (engine === 'DI' && kind === 'shape') {
        expect(h.context().geometry!.isPageBuilt(0)).toBe(true);
        const [text] = await h.context().geometry!.readAnchorGeometries([{ kind: 'range', version: h.context().snapshot.version,
          range: { story: 'body', start: { paraId: '00000001', offset: 0 }, end: { paraId: '00000001', offset: 6 }, view: 'accepted' } }]);
        expect(text).toMatchObject({ ok: true });
      }
      expect(h.settles).toHaveLength(1);
      expect(Object.keys(h.context().snapshot.selection).sort()).toEqual(['displayRange', 'formatting']);
      const selected = (await h.readAnchor())[0]!;
      const [range] = await h.context().geometry!.readAnchorGeometries([{ kind: 'range', version: h.context().snapshot.version,
        range: { story: 'body', start: { paraId: '00000001', offset: 7 }, end: { paraId: '00000001', offset: 8 }, view: 'accepted' } }]);
      if (engine === 'DI' && kind === 'shape' && !selected.ok) {
        expect(selected.failure.code, `Shape selection anchor failed: ${selected.failure.code}`).toBe('missing-target');
        return;
      }
      expectSelectionRange(selected, range!, h.context(), engine === 'DI' ? 'trailing' : 'caret');
      expect(selected.ok || selected.failure.code === 'missing-target').toBe(true);
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') test('T16/B12: a multi-cell selection exposes its text range or missing-target', async () => {
    const bytes = await storyAnchorsDocx();
    const h = await matrixEditor(engine, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    try {
      await h.selection('body:t0:r0c0', '00000002', 1);
      await h.quiet();
      h.reset();
      await act(async () => {
        if (h.controller()) expect(await h.controller()!.submit({ kind: 'selectTableCells', payload: { target: 'table' } })).toMatchObject({ ok: true });
        else expect(h.ref.current!.getEditorRef()!.applyYrsCommand({ type: 'tableSelect', target: 'table' })).toBe(true);
      });
      await h.quiet();
      expect(Object.keys(h.context().snapshot.selection).sort()).toEqual(['displayRange', 'formatting']);
      if (h.controller()) expect(h.controller()!.getSnapshot().view!.editState!.tableContext?.hasMultiCellSelection).toBe(true);
      else {
        const cells = h.session().cellSelection()!;
        expect(cells).not.toBeNull();
        expect(cells.anchor.column).not.toBe(cells.head.column);
      }
      const selected = (await h.readAnchor())[0]!;
      const raw = (h.controller() ? h.controller()!.getSnapshot().view!.workerSelection : h.session().selection())!;
      expect(raw).toBeTruthy();
      const focus = 'focus' in raw ? raw.focus : raw.head;
      expect(raw.anchor.story).toBe(focus.story);
      const [textRange] = await h.context().geometry!.readAnchorGeometries([{ kind: 'range', version: h.context().snapshot.version,
        range: { story: raw.anchor.story, start: { paraId: raw.anchor.paraId, offset: raw.anchor.offset },
          end: { paraId: focus.paraId, offset: focus.offset }, view: 'accepted' } }]);
      expectSelectionRange(selected, textRange!, h.context());
      expect(selected.ok || selected.failure.code === 'missing-target').toBe(true);
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') test('T17: clearing the selection settles once with missing-target', async () => {
    const h = await matrixEditor(engine);
    try {
      await h.selection('body', h.first.paraId, 0, 3);
      await h.quiet();
      h.reset();
      await h.clearSelection();
      await h.quiet();
      expect(h.settles).toHaveLength(1);
      expect(h.settles[0]!.anchor).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
      expect(await h.readAnchor()).toMatchObject([{ ok: false, failure: { code: 'missing-target' } }]);
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
    } finally { h.dispose(); }
  });

  if (engine !== 'DL') test('T23: continuous typing longer than the window settles once after the last key', async () => {
    const h = await matrixEditor(engine, documentBytes(), undefined, true);
    try {
      await h.selection('body', h.first.paraId, 0);
      await h.quiet();
      h.reset();
      const input = h.input();
      let lastKeyAt = 0;
      for (let index = 0; index < 8; index++) {
        await act(async () => {
          lastKeyAt = Date.now();
          fireEvent.keyDown(input, { key: 'x' });
          if (h.controller()) expect(await h.controller()!.submit({ kind: 'insertText', payload: { text: 'x' } })).toMatchObject({ ok: true });
          else h.ref.current!.getEditorRef()!.insertText('x');
        });
        await h.ref.current!.flushPendingInput();
        await h.quiet(100);
        expect(h.settles).toHaveLength(0);
      }
      expect(h.settles).toHaveLength(0);
      await h.selectionLayout();
      await h.quiet(249);
      expect(h.settles).toHaveLength(0);
      await h.quiet(1);
      expect(h.settles).toHaveLength(1);
      expect(h.settles[0]!.at - lastKeyAt).toBeGreaterThanOrEqual(350);
      expect(h.context().snapshot.selectionSettled).toBe(true);
      await h.quiet(700);
      expect(h.settles).toHaveLength(1);
    } finally { h.dispose(); }
  }, 15_000);
});


async function storedStoryDocument() {
  const body = '<w:p w14:paraId="00000001"><w:commentRangeStart w:id="1"/><w:r><w:t>Body text</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
    '<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="box" style="width:100pt;height:40pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Box label</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>';
  const zip = await JSZip.loadAsync(await inlineImageDocument(body));
  zip.file('word/comments.xml', '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Author"><w:p><w:r><w:t>Comment label</w:t></w:r></w:p></w:comment></w:comments>');
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
  zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>', '<Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
  return zip.generateAsync({ type: 'arraybuffer' });
}

for (const engine of ['DI', 'DL', 'DP', 'DV'] as const) test(`T15/B9: ${engine} comment and text-box text selections are unreachable today`, async () => {
  const buffer = await storedStoryDocument();
  const zip = await JSZip.loadAsync(buffer);
  expect(await zip.file('word/document.xml')!.async('string')).toContain('Box label');
  expect(await zip.file('word/comments.xml')!.async('string')).toContain('Comment label');
  const h = await matrixEditor(engine, buffer);
  try {
    const listed = await h.context().read.listStories();
    expect(listed.ok).toBe(true);
    if (!listed.ok) throw new Error('Story read failed');
    expect(listed.stories.some(story => story.kind === 'body')).toBe(true);
    expect(listed.stories.some(story => /comment|text.?box|txbx/i.test(story.story))).toBe(false);
    const read = await h.context().read.readParagraphs({ view: 'accepted' });
    expect(read.ok).toBe(true);
    if (!read.ok) throw new Error('Paragraph read failed');
    expect(read.paragraphs.map(paragraph => paragraph.text).join('\n')).toContain('Body text');
    expect(read.paragraphs.map(paragraph => paragraph.text).join('\n')).not.toMatch(/Box label|Comment label/);
  } finally { h.dispose(); }
});
