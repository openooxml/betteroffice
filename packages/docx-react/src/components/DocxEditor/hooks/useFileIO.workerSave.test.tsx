import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useCallback, useRef, useState } from 'react';
import JSZip from 'jszip';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import {
  createYrsSession,
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
import { registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
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

async function workerOpened(viewer = true) {
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
        renderedFrame: null,
        hydrateOnDemand: viewer,
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
    return { core, io, pagedEditorRef, ref, host, setHost, controller, workerDocument };
  }, { initialProps: 1 });
  await waitFor(() => expect(hook.result.current.core.session).not.toBeNull());
  const session = hook.result.current.core.session!;
  return { hook, session, worker: workers.at(-1)!, saved, errors, opens, flush, project };
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

test('an editor flushes its loaded peer and posts its diff immediately before save', async () => {
  const opened = await workerOpened(false);
  await act(async () => { await requestWorkerOpenReplica(opened.session); });
  opened.flush.mockImplementation(async () => {
    const first = opened.session.paragraphs('body')[0]!;
    opened.session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Peer edit ');
  });
  const posted = opened.worker.requests.length;
  const buffer = await opened.hook.result.current.io.handleSave();
  expect(buffer).toBeInstanceOf(ArrayBuffer);
  expect(opened.worker.requests.slice(posted)).toEqual(['applyUpdate', 'save']);
  const zip = await JSZip.loadAsync(buffer!);
  expect(await zip.file('word/document.xml')!.async('string')).toContain('Peer edit ');
  expect(opened.flush).toHaveBeenCalledTimes(1);
  expect(opened.project).not.toHaveBeenCalled();
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
