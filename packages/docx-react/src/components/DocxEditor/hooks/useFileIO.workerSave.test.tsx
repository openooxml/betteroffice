import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useCallback, useRef, useState } from 'react';
import JSZip from 'jszip';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import * as yrs from '@betteroffice/docx/yrs';
import {
  createYrsSession,
  ResidentEngineWorkerClient,
  ResidentWorkerSaveUnavailableError,
  type YrsDocxHost,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { residentWorkerFactory, type InProcessResidentWorker } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import { createStyleResolver } from '@betteroffice/docx/styles';
import { isMacPlatform } from '../../../commands/descriptors';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { createCommentIdAllocator } from '../commentFactories';
import { awaitWorkerOpenReplica, requestWorkerOpenReplica, workerOpenReplicaStarted } from '../internals/workerOpenReplica';
import { registerWorkerOpenSave, workerOpenSave } from '../internals/workerOpenSave';
import { registeredWorkerProposalAuthority, registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import { useCanvasRenderer, type OpenInWorker, type WorkerOpenedDocument } from './useDisplayList';
import { useDocxCommandBinding, type DocxCommandInputs } from './useDocxCommands';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';
import { useFileIO } from './useFileIO';
import { useYrsCoreSession } from './useYrsCoreSession';
import { useKeyboardShortcuts } from './useKeyboardShortcuts';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const ROOT = resolve(import.meta.dir, '../../../../../..');
const bytes = new Uint8Array(readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/page-fragments/pages.docx')));
const originalWorker = globalThis.Worker;
const sessions: YrsSession[] = [];
let startWorker!: () => InProcessResidentWorker;
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>>;
let workers: InProcessResidentWorker[];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))));
  startWorker = await residentWorkerFactory();
});
beforeEach(() => {
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  workers = [];
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
});
afterEach(() => {
  cleanup();
  compileModule.mockRestore();
  globalThis.Worker = originalWorker;
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function workerOpened(viewer = true, workerProposals = false) {
  let client!: ResidentEngineWorkerClient;
  const open = ResidentEngineWorkerClient.prototype.open;
  const capture = spyOn(ResidentEngineWorkerClient.prototype, 'open').mockImplementation(function (
    this: ResidentEngineWorkerClient, ...args: Parameters<ResidentEngineWorkerClient['open']>
  ) {
    client = this;
    return open.apply(this, args);
  });
  const saved: ArrayBuffer[] = [];
  const errors: Error[] = [];
  const opens: boolean[] = [];
  const flush = mock(async () => {});
  const project = mock(() => {});
  const hook = renderHook((generation: number) => {
    const renderer = useCanvasRenderer(undefined, undefined, undefined, undefined, undefined, true);
    const workerDocument = useRef<WorkerOpenedDocument | null>(null);
    const openInWorker = useCallback<OpenInWorker>(async (...args) => {
      workerDocument.current = await renderer.openInWorker(...args);
      return workerDocument.current;
    }, [renderer.openInWorker]);
    const [host, setHost] = useState<YrsDocxHost | null>(null);
    const [renderedFrame, setRenderedFrame] = useState<object | null>(null);
    const core = useYrsCoreSession(true, host?.document ?? null, null, bytes, generation, undefined, {
      onSession: (session) => {
        renderer.recordSession(session);
        const open = session.openDocx.bind(session);
        session.openDocx = (input, seed, options) => {
          opens.push(seed);
          return open(input, seed, options);
        };
      },
      onHostDocument: (document) => setHost(document),
      onError: (error) => errors.push(error),
    }, {
      workerOpen: {
        openInWorker,
        renderedFrame,
        pendingCompletion: renderer.pendingCompletion,
        workerProposals,
        hydrateOnDemand: viewer || workerProposals,
        onWorkerRevisions: () => {},
      },
    });
    const pagedEditorRef = useRef<PagedEditorRef | null>(null);
    pagedEditorRef.current = core.session ? {
      getYrsSession: () => core.session,
      isWorkerViewer: () => viewer,
      isFocused: () => true,
      flushPendingInput: async () => {
        await awaitWorkerOpenReplica(core.session!);
        await flush();
      },
      getDocument: () => {
        project();
        return core.documentFromYrs();
      },
    } as unknown as PagedEditorRef : null;
    const io = useFileIO({
      pagedEditorRef,
      viewerSession: viewer,
      resolveImage: () => null,
      comments: [],
      documentName: undefined,
      onSave: (buffer) => saved.push(buffer),
      downloadOnSave: false,
      onOpen: undefined,
      onError: (error) => errors.push(error),
      onPrint: undefined,
      onDocumentNameChange: undefined,
      loadBuffer: async () => {},
      focusActiveEditor: () => {},
    });
    const { controller } = useDocxCommandBinding({
      pagedEditorRef,
      session: core.session,
      document: host?.document ?? null,
      viewerSession: viewer,
      readOnly: viewer,
      mode: viewer ? 'viewing' : 'editing',
      experimentalWorkerOpen: true,
      bridgeRef: { current: {
        session: () => core.session,
        rootStory: () => 'body',
        hasPendingInput: () => false,
        hasSelection: () => false,
        toolbarSelection: () => null,
        selectedImage: () => null,
        subscribe: () => () => {},
        runAfterPendingInput: () => { throw new Error('Save must not wait in the input queue'); },
      } },
      save: io.handleDownloadDocument,
    } as unknown as DocxCommandInputs);
    useKeyboardShortcuts({
      commands: controller,
      pagedEditorRef,
      disableFindReplaceShortcuts: true,
      tableSelection: { state: { tableIndex: null } } as never,
    });
    const ref = useRef<DocxEditorRef>(null);
    useDocxEditorRefApi({
      ref,
      experimentalWorkerOpen: true,
      document: host?.document ?? null,
      documentFromYrs: core.documentFromYrs,
      historyStateRef: { current: host?.document ?? null },
      pagedEditorRef,
      handleSave: io.handleSave,
      zoom: 1,
      setZoom: () => {},
      scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
      loadParsedDocument: () => {},
      loadBuffer: async () => {},
      comments: [],
      setComments: () => {},
      setShowCommentsSidebar: () => {},
      contentChangeSubscribersRef: { current: new Set() },
      selectionChangeSubscribersRef: { current: new Set() },
      getCachedStyleResolver: createStyleResolver,
      commentIdAllocator: createCommentIdAllocator(),
      commands: controller.store,
      modeRef: { current: viewer ? 'viewing' : 'editing' },
      allowHostProposalsRef: { current: false },
      hostSearch: {
        search: async () => ({ query: '', options: { caseSensitive: false }, total: 0, current: -1 }),
        searchNext: () => null,
        searchPrevious: () => null,
        searchGoTo: () => null,
        clearSearch: () => {},
        getSearchState: () => null,
        onSearchChange: () => () => {},
      },
    });
    return { core, io, pagedEditorRef, ref, host, setHost, controller, workerDocument, setRenderedFrame };
  }, { initialProps: 1 });
  try {
    await waitFor(() => expect(hook.result.current.core.session).not.toBeNull());
  } finally {
    capture.mockRestore();
  }
  const session = hook.result.current.core.session!;
  return { hook, session, client, worker: workers.at(-1)!, saved, errors, opens, flush, project };
}

async function layOut(opened: Awaited<ReturnType<typeof workerOpened>>) {
  const metadata = await createYrsSession({});
  sessions.push(metadata);
  metadata.registerFont(new Uint8Array(readFileSync(resolve(ROOT, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))));
  metadata.adoptResidentWorkerLayout!(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
    renderEnv: {},
  }));
  await opened.client.bootstrap(
    { ...metadata.residentWorkerSnapshot()!, workerAuthoritative: true },
    '{}',
    { opened: true, layoutExtras: '{}' }
  );
}

test.each(['save', 'download', 'ref'] as const)('viewer %s saves in the worker without starting its replica', async (kind) => {
  const opened = await workerOpened();
  const { io, ref } = opened.hook.result.current;
  const buffer = kind === 'ref' ? await ref.current!.save()
    : kind === 'download' ? await io.handleDownloadDocument() : await io.handleSave();
  expect(kind === 'download' ? buffer : buffer instanceof ArrayBuffer).toBe(kind === 'download' ? 'saved' : true);
  expect(opened.worker.requests.filter((type) => type === 'save')).toHaveLength(1);
  expect(opened.worker.requests).not.toContain('encodeState');
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
  expect(opened.flush).not.toHaveBeenCalled();
  expect(opened.project).not.toHaveBeenCalled();
  expect(opened.opens).toEqual([]);
  expect(opened.saved).toHaveLength(1);
  expect(opened.errors).toEqual([]);
});

test.each(['command', 'shortcut'] as const)('viewer %s saves through the worker without starting the replica', async (kind) => {
  const opened = await workerOpened();
  if (kind === 'command') {
    expect(await opened.hook.result.current.controller.store.execute('save', null)).toEqual({ ok: true, status: 'executed' });
  } else {
    act(() => window.document.dispatchEvent(new KeyboardEvent('keydown', {
      key: 's', cancelable: true, ...(isMacPlatform() ? { metaKey: true } : { ctrlKey: true }),
    })));
    await waitFor(() => expect(opened.saved).toHaveLength(1));
  }
  expect(opened.worker.requests).toContain('save');
  expect(opened.worker.requests).not.toContain('encodeState');
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
  expect(opened.flush).not.toHaveBeenCalled();
  expect(opened.project).not.toHaveBeenCalled();
  expect(opened.errors).toEqual([]);
});

test('viewer save posts after a queued proposal call', async () => {
  const opened = await workerOpened();
  await layOut(opened);
  const authority = registerWorkerProposalAuthority(opened.session, opened.hook.result.current.workerDocument.current!, {
    current: () => true,
    laidOut: async () => {},
    relayout: () => {},
    adopted: () => {},
    handedOver: () => {},
    contentChanged: () => {},
  });
  await authority.initialize();
  const posted = opened.worker.requests.length;
  opened.worker.hold();
  const proposal = authority.propose({ expectVersion: opened.session.version(), proposals: [] }, async () => {
    throw new Error('Unexpected main proposal');
  });
  const saving = opened.hook.result.current.io.handleSave();
  await waitFor(() => expect(opened.worker.requests.slice(posted)).toEqual(['proposal']));
  opened.worker.release();
  await proposal;
  expect(await saving).toBeInstanceOf(ArrayBuffer);
  expect(opened.worker.requests.slice(posted)).toEqual(['proposal', 'save', 'proposal']);
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
  expect(opened.flush).not.toHaveBeenCalled();
  expect(opened.project).not.toHaveBeenCalled();
});

test('an editor with an unloaded peer saves without hydrating', async () => {
  const opened = await workerOpened(false);
  expect(await opened.hook.result.current.ref.current!.save()).toBeInstanceOf(ArrayBuffer);
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
  expect(opened.worker.requests).not.toContain('encodeState');
  expect(opened.flush).not.toHaveBeenCalled();
  expect(opened.project).not.toHaveBeenCalled();
  expect(opened.errors).toEqual([]);
});

test('a save without the editing copy clears the editor\'s save marks for the next editor save', async () => {
  const opened = await workerOpened(false, true);
  await layOut(opened);
  act(() => opened.hook.result.current.setRenderedFrame({}));
  await waitFor(() => expect(opened.hook.result.current.core.workerProposalsReady).toBe(true));
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
  expect(opened.hook.result.current.core.replicaReady).toBe(false);
  expect(opened.session.storyIds()).toEqual([]);
  expect(opened.opens).toEqual([]);
  const authority = registeredWorkerProposalAuthority(opened.session)!;
  const identities = await authority.paragraphIdentities(async () => {
    throw new Error('Unexpected main paragraph identities');
  });
  const story = 'hf:rIdHeader1';
  const paragraph = identities.paragraphs.find((identity) => identity.session?.story === story)!.session!;
  expect(paragraph.story).not.toBe('body');
  await act(async () => {
    expect(await authority.propose({
      expectVersion: opened.session.version(),
      proposals: [{
        id: 'header-proposal', paragraph,
        suggest: { author: 'Host', date: '2026-09-29T12:00:00Z' },
        op: 'insertText', at: 'start', text: 'Worker header ',
      }],
    }, async () => {
      throw new Error('Unexpected main proposal');
    })).toMatchObject({
      ok: true,
      snapshot: { proposals: [{ changed: true, paragraph: { story } }] },
    });
  });
  expect(authority.holdsCommittedWorkerState()).toBe(true);
  const save = spyOn(opened.hook.result.current.workerDocument.current!, 'save');
  try {
    expect(workerOpenReplicaStarted(opened.session)).toBe(false);
    expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
    expect(save.mock.calls).toHaveLength(1);
    expect(save.mock.calls[0]![0]).not.toHaveProperty('stories');
    expect(save.mock.calls[0]![1]).toBeUndefined();
    expect(workerOpenReplicaStarted(opened.session)).toBe(false);
    expect(opened.worker.requests).not.toContain('encodeState');
    expect(opened.flush).not.toHaveBeenCalled();
    await act(async () => { await requestWorkerOpenReplica(opened.session); });
    expect(opened.hook.result.current.core.replicaReady).toBe(true);
    expect(opened.opens).toEqual([false]);
    expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
    expect(save.mock.calls).toHaveLength(2);
    expect(save.mock.calls[1]![0].stories).toEqual([]);
    expect(save.mock.calls[1]![1]).toBe(opened.session);
    expect(opened.project).not.toHaveBeenCalled();
    expect(opened.errors).toEqual([]);
  } finally {
    save.mockRestore();
  }
});

test('an editor flushes its loaded peer and posts its diff immediately before save', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  opened.flush.mockImplementation(async () => {
    const first = opened.session.paragraphs('body')[0]!;
    opened.session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Peer edit ');
    opened.hook.result.current.core.publishDirectInput('body');
  });
  const posted = opened.worker.requests.length;
  const save = spyOn(opened.hook.result.current.workerDocument.current!, 'save');
  const buffer = await opened.hook.result.current.io.handleSave();
  expect(buffer).toBeInstanceOf(ArrayBuffer);
  expect(opened.worker.requests.slice(posted)).toEqual(['applyUpdate', 'save']);
  const zip = await JSZip.loadAsync(buffer!);
  expect(await zip.file('word/document.xml')!.async('string')).toContain('Peer edit ');
  expect(opened.flush).toHaveBeenCalledTimes(1);
  expect(opened.project).not.toHaveBeenCalled();
  expect(opened.errors).toEqual([]);
  expect(save.mock.calls[0]![0].stories).toEqual(['body']);
  opened.flush.mockImplementation(async () => {});
  await opened.hook.result.current.io.handleSave();
  expect(save.mock.calls[1]![0].stories).toEqual([]);
  save.mockRestore();
});

test('a worker save keeps body edits already projected by getDocument', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  const { core, pagedEditorRef, io, workerDocument } = opened.hook.result.current;
  const body = opened.session.paragraphs('body')[0]!;
  opened.session.insertText({ story: 'body', paraId: body.paraId, offset: 0 }, 'Body edit ');
  core.publishDirectInput('body');
  const projected = pagedEditorRef.current!.getDocument();
  expect(projected).not.toBeNull();
  expect(JSON.stringify(projected!.package.document.content)).toContain('Body edit ');
  const story = 'hf:rIdHeader1';
  const header = opened.session.paragraphs(story)[0]!;
  opened.session.insertText({ story, paraId: header.paraId, offset: 0 }, 'Header edit ');
  core.publishDirectInput(story);
  const save = spyOn(workerDocument.current!, 'save');
  try {
    const buffer = await io.handleSave();
    expect(buffer).toBeInstanceOf(ArrayBuffer);
    const zip = await JSZip.loadAsync(buffer!);
    expect(await zip.file('word/document.xml')!.async('string')).toContain('Body edit ');
    expect(await zip.file('word/header1.xml')!.async('string')).toContain('Header edit ');
    expect(save.mock.calls[0]![0].stories).toEqual(['body', story]);
    expect(opened.errors).toEqual([]);
  } finally {
    save.mockRestore();
  }
});

test.each(['getDocument', 'fallback'] as const)('%s keeps body edits after a worker save and a header edit', async (kind) => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  const { core, pagedEditorRef, io } = opened.hook.result.current;
  expect(pagedEditorRef.current!.getDocument()).not.toBeNull();
  const body = opened.session.paragraphs('body')[0]!;
  opened.session.insertText({ story: 'body', paraId: body.paraId, offset: 0 }, 'Body edit ');
  core.publishDirectInput('body');
  const saved = await io.handleSave();
  expect(saved).toBeInstanceOf(ArrayBuffer);
  const first = await JSZip.loadAsync(saved!);
  expect(await first.file('word/document.xml')!.async('string')).toContain('Body edit ');
  const story = 'hf:rIdHeader1';
  const header = opened.session.paragraphs(story)[0]!;
  opened.session.insertText({ story, paraId: header.paraId, offset: 0 }, 'Header edit ');
  core.publishDirectInput(story);
  if (kind === 'getDocument') {
    const projected = pagedEditorRef.current!.getDocument();
    expect(projected).not.toBeNull();
    expect(JSON.stringify(projected!.package.document.content)).toContain('Body edit ');
    expect(JSON.stringify(projected!.package.headers?.get('rIdHeader1')?.content)).toContain('Header edit ');
  } else {
    registerWorkerOpenSave(opened.session, {
      available: () => false,
      save: async () => { throw new Error('Unexpected worker save'); },
    });
    const buffer = await io.handleSave();
    expect(buffer).toBeInstanceOf(ArrayBuffer);
    const zip = await JSZip.loadAsync(buffer!);
    expect(await zip.file('word/document.xml')!.async('string')).toContain('Body edit ');
    expect(await zip.file('word/header1.xml')!.async('string')).toContain('Header edit ');
  }
  expect(opened.errors).toEqual([]);
});

test('a worker save after getDocument with no later edit projects every story', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  const { core, pagedEditorRef, io, workerDocument } = opened.hook.result.current;
  const body = opened.session.paragraphs('body')[0]!;
  opened.session.insertText({ story: 'body', paraId: body.paraId, offset: 0 }, 'Body edit ');
  core.publishDirectInput('body');
  expect(pagedEditorRef.current!.getDocument()).not.toBeNull();
  const save = spyOn(workerDocument.current!, 'save');
  try {
    const buffer = await io.handleSave();
    expect(save.mock.calls[0]![0].stories).toEqual([]);
    const zip = await JSZip.loadAsync(buffer!);
    expect(await zip.file('word/document.xml')!.async('string')).toContain('Body edit ');
    expect(opened.errors).toEqual([]);
  } finally {
    save.mockRestore();
  }
});

test('an unavailable worker save preserves peer stories for the main-thread fallback', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  opened.hook.result.current.core.publishDirectInput('body');
  const save = spyOn(opened.hook.result.current.workerDocument.current!, 'save')
    .mockRejectedValue(new ResidentWorkerSaveUnavailableError('No source package'));
  const project = spyOn(yrs, 'yrsToDocument');
  try {
    expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
    expect(save.mock.calls[0]![0].stories).toEqual(['body']);
    expect(project.mock.calls[0]![2]?.storyIds).toEqual(new Set(['body']));
    expect(opened.errors).toEqual([]);
  } finally {
    save.mockRestore();
    project.mockRestore();
  }
});

test('a worker save retains peer edits marked while its response is pending', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  const core = opened.hook.result.current.core;
  core.publishDirectInput('body');
  const save = spyOn(opened.hook.result.current.workerDocument.current!, 'save');
  const saver = workerOpenSave(opened.session)!;
  opened.worker.hold();
  const pending = saver.save([], opened.session);
  try {
    await waitFor(() => expect(opened.worker.requests).toContain('save'));
    const paragraph = opened.session.paragraphs('body')[0]!;
    opened.session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Later edit ');
    core.publishDirectInput('body');
    opened.worker.release();
    await pending;
    const buffer = await saver.save([], opened.session);
    expect(save.mock.calls[1]![0].stories).toEqual(['body']);
    expect(await (await JSZip.loadAsync(buffer)).file('word/document.xml')!.async('string')).toContain('Later edit ');
    await saver.save([], opened.session);
    expect(save.mock.calls[2]![0].stories).toEqual([]);
  } finally {
    opened.worker.release();
    save.mockRestore();
  }
});

test('an editor integrates saved paragraph ID claims before later worker proposals', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  await layOut(opened);
  let paraId!: string;
  let beforeSave!: Uint8Array;
  const origins: string[] = [];
  let unsubscribe = () => {};
  opened.flush.mockImplementation(async () => {
    const first = opened.session.paragraphs('body')[0]!;
    paraId = opened.session.splitParagraph({ story: 'body', paraId: first.paraId, offset: 0 }).secondParaId;
    opened.session.insertText({ story: 'body', paraId, offset: 0 }, 'Peer claim ');
    opened.hook.result.current.core.publishDirectInput('body');
    beforeSave = opened.session.encodeStateVector();
    unsubscribe = opened.session.onUpdate((_update, origin) => origins.push(origin));
  });
  try {
    await act(async () => {
      expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
    });
  } finally {
    unsubscribe();
  }
  expect(origins.length).toBeGreaterThan(0);
  expect(new Set(origins)).toEqual(new Set(['remote']));
  const saveIndex = opened.worker.requests.lastIndexOf('save');
  expect(saveIndex).toBeGreaterThanOrEqual(0);
  expect(opened.worker.requests.slice(saveIndex + 1)).not.toContain('applyUpdate');
  expect(opened.session.encodeStateVector()).not.toEqual(beforeSave);
  expect(opened.session.encodeStateVector()).toEqual(opened.worker.sessions[0]!.encodeStateVector());
  const paragraph = opened.session.paragraphIdentities().paragraphs
    .find((identity) => identity.session?.paraId === paraId)!.persisted!;
  const initial = await opened.client.proposal({ kind: 'snapshot' });
  const proposed = await opened.client.proposal({
    kind: 'propose',
    request: {
      expectVersion: initial.mirror.version,
      proposals: [{
        id: 'after-save', paragraph,
        suggest: { author: 'Host', date: '2026-09-29T12:00:00Z' },
        op: 'replaceText', search: 'Peer claim ', replaceWith: 'Worker proposal ',
      }],
    },
  });
  expect(proposed.result?.ok).toBe(true);
  expect(proposed.updates.length).toBeGreaterThan(0);
  await act(async () => {
    for (const update of proposed.updates) opened.session.applyUpdate(update);
  });
  expect(opened.session.encodeStateVector()).toEqual(opened.worker.sessions[0]!.encodeStateVector());
  expect(opened.session.readParagraphs({ story: 'body', paraIds: [paraId], view: 'accepted' }))
    .toMatchObject({ ok: true, paragraphs: [{ text: expect.stringContaining('Worker proposal ') }] });
  expect(opened.errors).toEqual([]);
});

test('a loaded peer with an empty diff posts only save', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  const posted = opened.worker.requests.length;
  expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
  expect(opened.worker.requests.slice(posted)).toEqual(['save']);
});

test('an editor waits for an in-flight peer load before flushing and saving', async () => {
  const opened = await workerOpened(false);
  opened.worker.hold();
  const ready = requestWorkerOpenReplica(opened.session);
  await waitFor(() => expect(opened.worker.requests).toContain('encodeState'));
  const saving = opened.hook.result.current.io.handleSave();
  expect(opened.worker.requests).not.toContain('save');
  expect(opened.flush).not.toHaveBeenCalled();
  await act(async () => {
    opened.worker.release();
    await ready;
    expect(await saving).toBeInstanceOf(ArrayBuffer);
  });
  expect(opened.flush).toHaveBeenCalledTimes(1);
  expect(opened.project).not.toHaveBeenCalled();
  expect(opened.errors).toEqual([]);
});

test.each([true, false])('worker save unavailable with viewer=%s (a viewer reports it and never requests its copy)', async (viewer) => {
  const opened = await workerOpened(viewer);
  registerWorkerOpenSave(opened.session, {
    available: () => true,
    save: async () => { throw new ResidentWorkerSaveUnavailableError('No source package'); },
  });
  const buffer = await opened.hook.result.current.io.handleSave();
  if (viewer) {
    expect(buffer).toBeNull();
    expect(opened.errors).toEqual([expect.any(ResidentWorkerSaveUnavailableError)]);
    expect(opened.saved).toEqual([]);
    expect(opened.project).not.toHaveBeenCalled();
    expect(opened.opens).toEqual([]);
    expect(workerOpenReplicaStarted(opened.session)).toBe(false);
  } else {
    expect(buffer).toBeInstanceOf(ArrayBuffer);
    expect(opened.errors).toEqual([]);
    expect(opened.opens).toEqual([false]);
    expect(opened.project).toHaveBeenCalledTimes(1);
    expect(opened.saved).toEqual([buffer!]);
  }
});

test('an editor without a live worker uses its main-thread peer', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  registerWorkerOpenSave(opened.session, {
    available: () => false,
    save: async () => { throw new Error('Unexpected worker save'); },
  });
  expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
  expect(opened.worker.requests).not.toContain('save');
  expect(opened.project).toHaveBeenCalledTimes(1);
  expect(opened.errors).toEqual([]);
});

test('a viewer with a main-thread selection surface still refuses a missing worker', async () => {
  const opened = await workerOpened();
  opened.hook.result.current.pagedEditorRef.current!.isWorkerViewer = () => false;
  registerWorkerOpenSave(opened.session, {
    available: () => false,
    save: async () => { throw new Error('No document worker'); },
  });
  expect(await opened.hook.result.current.io.handleSave()).toBeNull();
  expect(opened.errors.map((error) => error.message)).toEqual(['No document worker']);
  expect(opened.project).not.toHaveBeenCalled();
  expect(opened.flush).not.toHaveBeenCalled();
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
});

test('worker saves read current host metadata without projecting or changing the host package', async () => {
  const opened = await workerOpened();
  const { host, setHost, workerDocument } = opened.hook.result.current;
  const save = spyOn(workerDocument.current!, 'save');
  const currentHost = { ...host!, document: { ...host!.document, warnings: ['Current host metadata'] } };
  act(() => setHost(currentHost));
  const original = currentHost.document.originalBuffer;
  expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
  const request = save.mock.calls[0]![0];
  expect(request.host?.warnings).toEqual(['Current host metadata']);
  expect(request.host?.originalBuffer).toBeUndefined();
  expect(request.host?.package.document.content).toEqual([]);
  expect(currentHost.document.originalBuffer).toBe(original);
  expect(opened.project).not.toHaveBeenCalled();
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
  save.mockRestore();
});

test('a terminal viewer worker reports the error without hydrating', async () => {
  const opened = await workerOpened();
  opened.worker.onerror?.({ message: 'Worker stopped', preventDefault() {} } as ErrorEvent);
  expect(await opened.hook.result.current.io.handleSave()).toBeNull();
  expect(opened.errors).toHaveLength(1);
  expect(opened.opens).toEqual([]);
  expect(opened.project).not.toHaveBeenCalled();
  expect(workerOpenReplicaStarted(opened.session)).toBe(false);
});

test('a document replaced mid-save reports the replacement and discards the bytes', async () => {
  const opened = await workerOpened();
  opened.worker.hold();
  const pending = opened.hook.result.current.io.handleSave();
  await waitFor(() => expect(opened.worker.requests).toContain('save'));
  const other = await createYrsSession();
  sessions.push(other);
  opened.hook.result.current.pagedEditorRef.current = { getYrsSession: () => other } as PagedEditorRef;
  opened.worker.release();
  expect(await pending).toBeNull();
  expect(opened.errors.map((error) => error.message)).toEqual(['The document changed while saving']);
  expect(opened.saved).toEqual([]);
});

test('unmount clears the session worker saver', async () => {
  const opened = await workerOpened();
  expect(workerOpenSave(opened.session)).not.toBeNull();
  opened.hook.unmount();
  expect(workerOpenSave(opened.session)).toBeNull();
});

test('replacing the document clears its saver and saves through the new owner', async () => {
  const opened = await workerOpened();
  opened.hook.rerender(2);
  expect(workerOpenSave(opened.session)).toBeNull();
  await waitFor(() => {
    expect(opened.hook.result.current.core.session).not.toBeNull();
    expect(opened.hook.result.current.core.session).not.toBe(opened.session);
  });
  const saves = opened.worker.requests.filter((type) => type === 'save').length;
  expect(await opened.hook.result.current.io.handleSave()).toBeInstanceOf(ArrayBuffer);
  expect(opened.worker.requests.filter((type) => type === 'save')).toHaveLength(saves);
  expect(workers.at(-1)!.requests).toContain('save');
  expect(workerOpenReplicaStarted(opened.hook.result.current.core.session!)).toBe(false);
});
