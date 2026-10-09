import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from './PagedEditor';
import { DocxWorkerError } from './internals/docxWorkerError';
import { workerOpenDocumentHeld } from './internals/workerOpenReplica';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from './internals/engineChoice';
import * as replicaHelpers from './internals/workerOpenReplica';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import { residentWorkerFactory } from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import type { ResidentEngineWorkerRequest, ResidentEngineWorkerResponse } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

// The second session a load creates is the full one, after its preview.
const real = await import('@betteroffice/docx/yrs');
// Mocking rebinds the module's live exports, `real`'s included.
const { createYrsSession, ResidentWorkerOutOfMemoryError } = real;
let created = 0;
let fullSession: unknown = null;
let shownPages = false;
let fullOpen: 'fail' | 'open' | 'layout-fail' = 'fail';
let holdFullOpen: Promise<void> | null = null;
let failPreviewLayout = false;
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: async (options: Parameters<typeof createYrsSession>[0]) => {
    created += 1;
    if (created === 2 && holdFullOpen) await holdFullOpen;
    if (created === 2 && fullOpen === 'fail') {
      // Fails once the preview shows, however long a loaded machine takes to paint it.
      for (let waited = 0; !shownPages && waited < 20_000; waited += 20) {
        await new Promise((done) => setTimeout(done, 20));
      }
      throw new Error('full open failed');
    }
    const session = await createYrsSession(options);
    if (created === 1 && failPreviewLayout) {
      session.layoutFontRequirementsJson = () => {
        throw new Error('preview layout failed');
      };
    }
    if (created === 2) fullSession = session;
    if (created === 2 && fullOpen === 'layout-fail') {
      session.layoutFontRequirementsJson = () => {
        throw new Error('layout failed');
      };
    }
    return session;
  },
}));
const layoutRender = await import('@betteroffice/docx/layout/render');
const { buildRustDisplayFrame } = layoutRender;
let holdFullDisplayList: Promise<void> | null = null;
mock.module('@betteroffice/docx/layout/render', () => ({
  ...layoutRender,
  buildRustDisplayFrame: async (...args: Parameters<typeof buildRustDisplayFrame>) => {
    const result = await buildRustDisplayFrame(...args);
    if (holdFullDisplayList && args[1] === fullSession) await holdFullDisplayList;
    return result;
  },
}));
const displayList = await import('./hooks/useDisplayList');
const { useCanvasRenderer } = displayList;
let renderer: ReturnType<typeof useCanvasRenderer> | null = null;
let workerOpen: ReturnType<typeof useCanvasRenderer>['openInWorker'] | null = null;
let residentOpen: ReturnType<typeof useCanvasRenderer>['openInWorker'] | null = null;
let workerFailure: { error: Error; errorEngine: YrsSession } | null = null;
let holdCanvasReplay: 'full' | 'preview' | 'ordinary' | null = null;
interface PendingCanvasReplay {
  displayList: NonNullable<ReturnType<typeof useCanvasRenderer>['displayList']>;
  layoutEngine: unknown;
  isCurrent: () => boolean;
  resolve: () => void;
  reject: (error: Error) => void;
}
let pendingCanvasReplays: PendingCanvasReplay[] = [];
// Once the full session exists, its pages fail to render; or, until the full
// session's first frame shows, the preview's error stays set.
const renderFailure = new Error('render failed');
const previewFailure = new Error('preview render failed');
let failRender: 'full' | 'preview' | null = null;
mock.module('./hooks/useDisplayList', () => ({
  ...displayList,
  useCanvasRenderer: (...args: Parameters<typeof useCanvasRenderer>) => {
    renderer = useCanvasRenderer(...args);
    residentOpen = renderer.openInWorker;
    if (renderer.displayList) shownPages = true;
    if (workerOpen) renderer = { ...renderer, openInWorker: workerOpen };
    if (workerFailure) return { ...renderer, ...workerFailure, status: 'error' as const };
    const error =
      failRender === 'full' && created >= 2
        ? renderFailure
        : failRender === 'preview' &&
            renderer.presentedEngine &&
            (created < 2 || renderer.presentedEngine !== fullSession)
          ? previewFailure
          : null;
    if (error) {
      return {
        ...renderer,
        error,
        errorEngine: failRender === 'preview' ? renderer.presentedEngine : renderer.errorEngine,
        status: 'error' as const,
      };
    }
    return holdCanvasReplay ? { ...renderer, offscreenReplay: null } : renderer;
  },
}));
const canvasReplay = await import('./canvasReplay');
const { presentCanvasReplay } = canvasReplay;
mock.module('./canvasReplay', () => ({
  ...canvasReplay,
  presentCanvasReplay: (...args: Parameters<typeof presentCanvasReplay>) => {
    const [preparations, isCurrent] = args;
    const shown = renderer;
    if (
      !holdCanvasReplay ||
      !shown?.displayList ||
      (holdCanvasReplay === 'full' &&
        (fullSession === null || shown.presentedEngine !== fullSession)) ||
      (holdCanvasReplay === 'preview' && shown.presentedEngine === fullSession)
    ) {
      return presentCanvasReplay(...args);
    }
    const list = shown.displayList;
    const ready = new Promise<void>((resolve, reject) => {
      pendingCanvasReplays.push({
        displayList: list,
        layoutEngine: shown.layoutEngine,
        isCurrent,
        resolve,
        reject,
      });
    });
    return presentCanvasReplay(
      [...preparations, { buffer: document.createElement('canvas'), ready, present: () => {} }],
      isCurrent
    );
  },
}));
const { isPresented, onReplayFailed } = await import('./internals/layoutProvenance');
const coreSessionModule = await import('./hooks/useYrsCoreSession');
const { useYrsCoreSession } = coreSessionModule;
let replicaError: ((error: Error) => void) | null = null;
mock.module('./hooks/useYrsCoreSession', () => ({
  ...coreSessionModule,
  useYrsCoreSession: (...args: Parameters<typeof useYrsCoreSession>) => {
    replicaError = (error) => args[6]?.onReplicaError?.(error, args[4]);
    return useYrsCoreSession(...args);
  },
}));
const refApiModule = await import('./hooks/useDocxEditorRefApi');
const { useDocxEditorRefApi, DOCX_REF_REPLICA_LOADING_MUTATIONS, DocxReplicaNotReadyError } = refApiModule;
let hostEditorRef: React.RefObject<PagedEditorRef | null> | null = null;
mock.module('./hooks/useDocxEditorRefApi', () => ({
  ...refApiModule,
  useDocxEditorRefApi: (...args: Parameters<typeof useDocxEditorRefApi>) => {
    hostEditorRef = args[0].pagedEditorRef;
    return useDocxEditorRefApi(...args);
  },
}));
const { DocxEditor } = await import('../../index');
const { DocxAsyncOnlyError } = await import('./hooks/useDocxEditorRefApi');
type Editor = import('../../index').DocxEditorRef;

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const PAGES = resolve(
  import.meta.dir,
  '../../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx'
);
const quiet = { error: console.error, warn: console.warn };
const originalWorker = globalThis.Worker;

beforeEach(() => setMissingWorkerCapabilitiesForTests([]));
beforeAll(async () => {
  // Frames build on this thread, where the tests hold them; a resident worker's would not wait.
  globalThis.Worker = undefined as unknown as typeof Worker;
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
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  console.error = () => {};
  console.warn = () => {};
});
afterEach(() => {
  cleanup();
  resetEngineChoiceForTests();
  workerOpen = null;
  workerFailure = null;
  holdCanvasReplay = null;
  for (const replay of pendingCanvasReplays) replay.resolve();
  pendingCanvasReplays = [];
});
afterAll(async () => {
  // React's scheduler still runs the last commit's passive effects, which read `window`.
  await new Promise((done) => setTimeout(done, 100));
  console.error = quiet.error;
  console.warn = quiet.warn;
  globalThis.Worker = originalWorker;
  if (ownsDom) await GlobalRegistrator.unregister();
});

const documentBuffer = () => {
  const bytes = readFileSync(PAGES);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};
const load = (
  buffer: ArrayBuffer,
  onError: (error: Error) => void,
  ref = createRef<Editor>(),
  experimentalWorkerOpen = false
) => (
  <DocxEditor
    ref={ref}
    previewFirstPage
    experimentalWorkerOpen={experimentalWorkerOpen}
    documentBuffer={buffer}
    onError={onError}
  />
);

/** A wait begun after a load failed rejects with it. */
async function expectWaitRejects(ref: React.RefObject<Editor | null>) {
  let outcome = null as number | Error | null;
  ref.current!.whenLayoutComplete().then(
    (pages) => (outcome = pages),
    (error: Error) => (outcome = error)
  );
  await waitFor(() => expect(outcome).toBeInstanceOf(Error));
}

async function currentCanvasReplay() {
  await waitFor(() => expect(pendingCanvasReplays.some((replay) => replay.isCurrent())).toBe(true), {
    timeout: 20_000,
  });
  return pendingCanvasReplays.find((replay) => replay.isCurrent())!;
}

test('a load whose full open fails after its preview painted keeps none of its pages', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'fail';
  failRender = null;
  const bytes = readFileSync(PAGES);
  const ref = createRef<Editor>();
  const errors: string[] = [];
  const view = render(
    <DocxEditor
      ref={ref}
      previewFirstPage
      experimentalWorkerOpen={false}
      documentBuffer={
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
      }
      onError={(error) => errors.push(error.message)}
    />
  );
  await waitFor(() => expect(shownPages).toBe(true), { timeout: 10_000 });
  await waitFor(() => expect(errors).toEqual(['full open failed']), { timeout: 10_000 });
  await act(async () => {});
  expect(view.container.querySelector('.docx-editor-error')).not.toBeNull();
  expect(renderer!.displayList).toBeNull();
  expect(renderer!.presentedEngine).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
  await expectWaitRejects(ref);
}, 30_000);

test('an untaken worker session whose render and open fail reports the error once', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  const failure = new ResidentWorkerOutOfMemoryError('worker open ran out of memory', []);
  let release = () => {};
  const opening = new Promise<void>((done) => (release = done));
  const openInWorker = mock(async (_session: YrsSession) => {
    await opening;
    throw failure;
  });
  workerOpen = openInWorker;
  try {
    const ref = createRef<Editor>();
    const errors: Error[] = [];
    const buffer = documentBuffer();
    const onError = (error: Error) => errors.push(error);
    const view = render(load(buffer, onError, ref, true));
    await waitFor(() => expect(openInWorker).toHaveBeenCalledTimes(1), { timeout: 10_000 });
    const pendingSession = openInWorker.mock.calls[0][0];
    await waitFor(() => expect(shownPages).toBe(true), { timeout: 10_000 });
    expect(pendingSession as unknown).toBe(fullSession);
    expect(pendingSession.isDisplayOnly()).toBe(false);
    expect(renderer!.layoutEngine).not.toBe(pendingSession);
    expect((renderer!.layoutEngine as YrsSession).isDisplayOnly()).toBe(true);
    expect(ref.current!.getDocument()).toBeNull();

    workerFailure = { error: failure, errorEngine: pendingSession };
    await act(async () => view.rerender(load(buffer, onError, ref, true)));
    expect(errors).toEqual([]);

    await act(async () => release());
    await waitFor(() => expect(errors).toEqual([failure]), { timeout: 10_000 });
    await act(async () => {
      await new Promise((done) => setTimeout(done, 200));
    });
    expect(errors).toEqual([failure]);
    expect(openInWorker).toHaveBeenCalledTimes(1);
    expect(view.container.querySelector('.docx-editor-error')?.textContent).toContain(
      failure.message
    );
    await expectWaitRejects(ref);
  } finally {
    release();
  }
}, 30_000);

test.each([false, true])('a terminal worker open reports one typed error and an alert with viewer=%s', async (viewer) => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  const failure = new DocxWorkerError('open', new Error('worker open failed'));
  let release = () => {};
  const opening = new Promise<void>((done) => { release = done; });
  const openInWorker = mock(async (_session: YrsSession) => {
    await opening;
    throw failure;
  });
  workerOpen = openInWorker;
  const errors: Error[] = [];
  const buffer = documentBuffer();
  const element = () => (
    <DocxEditor
      experimentalWorkerOpen
      previewFirstPage={false}
      readOnly={viewer}
      documentBuffer={buffer}
      onError={(error) => errors.push(error)}
    />
  );
  const view = render(element());
  try {
    await waitFor(() => expect(openInWorker).toHaveBeenCalledTimes(1));
    const pending = openInWorker.mock.calls[0][0];
    workerFailure = { error: failure, errorEngine: pending };
    await act(async () => view.rerender(element()));
    expect(errors).toEqual([]);
    await act(async () => release());
    await waitFor(() => expect(errors).toEqual([failure]));
    expect(errors[0]).toBe(failure);
    expect(view.getByRole('alert').textContent).toContain(failure.message);
    expect(view.container.querySelector('.canvas-pages')).toBeNull();
    await act(async () => view.rerender(element()));
    expect(errors).toEqual([failure]);
    expect(openInWorker).toHaveBeenCalledTimes(1);
  } finally {
    release();
  }
});

test('a stale worker open failure keeps the replacement preview accepting input', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  const failure = new ResidentWorkerOutOfMemoryError('worker open ran out of memory', []);
  let releaseA = () => {};
  let releaseB = () => {};
  const openingA = new Promise<void>((done) => (releaseA = done));
  const openingB = new Promise<void>((done) => (releaseB = done));
  let opens = 0;
  let completeReplacement!: (opened: NonNullable<Awaited<ReturnType<NonNullable<typeof residentOpen>>>>) => void;
  const replacementOpened = new Promise<NonNullable<Awaited<ReturnType<NonNullable<typeof residentOpen>>>>>((resolve) => { completeReplacement = resolve; });
  const openInWorker = mock(async (...args: Parameters<NonNullable<typeof residentOpen>>) => {
    const first = ++opens === 1;
    await (first ? openingA : openingB);
    if (first) throw failure;
    const opened = await residentOpen!(...args);
    expect(opened).not.toBeNull();
    completeReplacement(opened!);
    return opened;
  });
  workerOpen = openInWorker;
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const buffer = documentBuffer();
  const onError = (error: Error) => errors.push(error);
  const view = render(load(buffer, onError, ref, true));
  const previousWorker = globalThis.Worker;
  let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>> | undefined;
  let releaseReplica = () => {};
  let restoreGate = () => {};
  try {
    await waitFor(() => expect(openInWorker).toHaveBeenCalledTimes(1), { timeout: 10_000 });
    const previousPreview = renderer!.layoutEngine;

    created = 0;
    fullSession = null;
    shownPages = false;
    await act(async () => view.rerender(load(buffer.slice(0), onError, ref, true)));
    await waitFor(
      () => {
        expect(openInWorker).toHaveBeenCalledTimes(2);
        expect(renderer!.layoutEngine).not.toBeNull();
        expect(renderer!.layoutEngine).not.toBe(previousPreview);
        expect(renderer!.presentedEngine).toBe(renderer!.layoutEngine);
        expect(renderer!.displayList).not.toBeNull();
        expect(isPresented(renderer!.canvasHostRef.current, renderer!.displayList!)).toBe(true);
      },
      { timeout: 10_000 }
    );
    const preview = renderer!.layoutEngine as YrsSession;
    expect(created).toBe(2);
    expect(openInWorker.mock.calls[1][0] as unknown).toBe(fullSession);
    expect(preview).not.toBe(fullSession);
    expect(preview.isDisplayOnly()).toBe(true);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false);
    expect(ref.current!.getDocument()).toBeNull();

    await act(async () => {
      releaseA();
      await new Promise((done) => setTimeout(done, 200));
    });
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false);
    expect(ref.current!.getDocument()).toBeNull();
    expect(errors).toEqual([]);
    expect(renderer!.layoutEngine).toBe(preview);
    expect(renderer!.presentedEngine).toBe(preview);
    expect(openInWorker).toHaveBeenCalledTimes(2);

    const startWorker = await residentWorkerFactory();
    compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
      new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
    ));
    let observedState!: () => void;
    const stateRequested = new Promise<void>((resolve) => { observedState = resolve; });
    const held: Array<() => void> = [];
    let released = false;
    releaseReplica = () => {
      released = true;
      for (const send of held.splice(0)) send();
    };
    globalThis.Worker = class {
      constructor() {
        const worker = startWorker();
        const send = worker.postMessage.bind(worker);
        worker.postMessage = (request, transfer) => {
          if (request.type === 'encodeState' && !released) {
            held.push(() => send(request, transfer));
            observedState();
          } else send(request, transfer);
        };
        return worker;
      }
    } as unknown as typeof Worker;
    const queuedText = 'Queued replacement ';
    fireEvent.input(view.getByTestId('yrs-input'), { target: { value: queuedText } });
    await act(async () => releaseB());
    const opened = await replacementOpened;
    await stateRequested;
    const session = fullSession as YrsSession;
    await waitFor(() => expect(hostEditorRef!.current!.getYrsSession()).toBe(session));
    const editor = hostEditorRef!.current!;
    const request = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
    const ensure = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
    const navigate = spyOn(editor, 'scrollToParaId');
    restoreGate = () => { request.mockRestore(); ensure.mockRestore(); navigate.mockRestore(); };
    const started = replicaHelpers.workerOpenReplicaStarted(session);
    const pendingApi = ref.current!;
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
    expect(pendingApi.getDocument()).toBeNull();
    expect(pendingApi.getEditorRef()).toBeNull();
    expect(pendingApi.findInDocument('Page')).toEqual([]);
    expect(pendingApi.getPageContent(1)).toBeNull();
    expect(pendingApi.getSelectionInfo()).toBeNull();
    expect(pendingApi.scrollToParaId('00000001')).toBe(false);
    const calls: Partial<Record<keyof Editor, unknown[]>> = {
      addComment: [{ paraId: '00000001', search: 'map', text: 'Check', author: 'Ann' }],
      proposeChange: [{ paraId: '00000001', search: 'map', replaceWith: 'plan', author: 'Host' }],
      applyFormatting: [{ paraId: '00000001', search: 'map', marks: { italic: true } }],
      setParagraphStyle: [{ paraId: '00000001', styleId: 'Normal' }],
      insertBreak: [{ paraId: '00000001', type: 'page' }],
    };
    for (const member of DOCX_REF_REPLICA_LOADING_MUTATIONS) {
      const args = calls[member];
      if (!args) throw new Error(`Missing loading mutation: ${member}`);
      expect(() => Reflect.apply(pendingApi[member] as Function, pendingApi, args)).toThrow(DocxReplicaNotReadyError);
    }
    expect(navigate).not.toHaveBeenCalled();
    expect(request).not.toHaveBeenCalled();
    expect(ensure).not.toHaveBeenCalled();
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(started);
    expect(editor.hasPendingInput()).toBe(true);
    const before = await opened.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
    expect(before.value.ok).toBe(true);
    if (!before.value.ok) throw new Error('The replacement worker could not read its document');
    const first = before.value.paragraphs[0]!;
    expect(first.paraId).toBe('00000001');
    expect(first.text).not.toContain(queuedText);
    request.mockRestore();
    ensure.mockRestore();
    await act(async () => releaseReplica());
    await waitFor(() => expect(ref.current!.getEditorRef()).not.toBeNull());
    await act(async () => ref.current!.flushPendingInput());
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(false);
    expect(ref.current!.getEditorRef()!.getYrsSession()).toBe(session);
    expect(ref.current!.getDocument()).not.toBeNull();
    expect(session.paragraphs('body')[0]!.text).toBe(`${queuedText}${first.text}`);
    expect(ref.current!.getEditorRef()!.hasPendingInput()).toBe(false);
    expect((await ref.current!.readParagraphs({ view: 'accepted' }))).toMatchObject({
      ok: true, paragraphs: [expect.objectContaining({ paraId: first.paraId, text: `${queuedText}${first.text}` }), ...before.value.paragraphs.slice(1)],
    });
    await act(async () => ref.current!.flushPendingInput());
    expect(session.paragraphs('body')[0]!.text).toBe(`${queuedText}${first.text}`);
    expect(ref.current!.findInDocument('Page')).toContainEqual({
      paraId: first.paraId, match: 'Page', before: queuedText, after: first.text.slice(4),
    });
    const readyEditor = ref.current!.getEditorRef()!;
    const readyLayout = spyOn(readyEditor, 'getLayout').mockReturnValue({
      pages: [{ fragments: [{ kind: 'paragraph', pmStart: 0 }] }],
    } as unknown as ReturnType<PagedEditorRef['getLayout']>);
    const readyLocation = spyOn(readyEditor, 'displayPositionToYrsLoc').mockReturnValue({
      story: 'body', paraId: first.paraId, offset: 0,
    });
    try {
      expect(ref.current!.getPageContent(1)).toMatchObject({
        pageNumber: 1,
        text: `[${first.paraId}] ${queuedText}${first.text}`,
        paragraphs: [{ paraId: first.paraId, text: `${queuedText}${first.text}` }],
      });
    } finally {
      readyLocation.mockRestore();
      readyLayout.mockRestore();
    }
    act(() => session.setSelection(
      { story: 'body', paraId: first.paraId, offset: queuedText.length },
      { story: 'body', paraId: first.paraId, offset: queuedText.length + 4 }
    ));
    expect(ref.current!.getSelectionInfo()).toMatchObject({
      paraId: first.paraId, selectedText: 'Page', paragraphText: `${queuedText}${first.text}`,
    });
    const target = { story: 'body', paraId: first.paraId, offset: 0 };
    act(() => session.setSelection({ ...target, offset: queuedText.length }));
    expect(ref.current!.scrollToParaId(first.paraId)).toBe(true);
    expect(session.selection()).toEqual({ anchor: target, head: target });
    const breaks = session.storySegments('body').filter((entry) => entry.kind === 'embed' && entry.embedKind === 'pageBreak').length;
    act(() => {
      for (const member of ['addComment', 'applyFormatting', 'setParagraphStyle', 'proposeChange', 'insertBreak'] as const) {
        const result = Reflect.apply(ref.current![member], ref.current, calls[member]!);
        if (member === 'addComment') {
          expect(result).toEqual(expect.any(Number));
          expect(session.resolveComment(String(result))).toContainEqual(expect.objectContaining({ story: 'body' }));
        } else expect(result).toBe(true);
        if (member === 'applyFormatting') {
          expect(session.storySegments('body')).toContainEqual(expect.objectContaining({
            kind: 'text', text: 'map', attributes: expect.objectContaining({ italic: true }),
          }));
        } else if (member === 'setParagraphStyle') {
          expect(session.paragraphs('body')[0]!.properties.pStyle).toBe('Normal');
        } else if (member === 'proposeChange') {
          expect(session.listRevisions()).toContainEqual(expect.objectContaining({ kind: 'insertion', author: 'Host', preview: 'plan' }));
        } else if (member === 'insertBreak') {
          expect(session.storySegments('body').filter((entry) => entry.kind === 'embed' && entry.embedKind === 'pageBreak')).toHaveLength(breaks + 1);
        }
      }
    });
    expect(() => session.yrsBlocksForStory('body')).not.toThrow();
    await act(async () => ref.current!.flushPendingInput());
    expect(session.paragraphs('body')[0]!.text.split(queuedText)).toHaveLength(2);
    expect(errors).toEqual([]);
  } finally {
    view.unmount();
    releaseA();
    releaseB();
    releaseReplica();
    restoreGate();
    compileModule?.mockRestore();
    globalThis.Worker = previousWorker;
  }
}, 30_000);

test('a load whose full session fails to render fails, and leaves no session behind', async () => {
  created = 0;
  fullOpen = 'open';
  failRender = 'full';
  const bytes = readFileSync(PAGES);
  const ref = createRef<Editor>();
  const errors: string[] = [];
  const view = render(
    <DocxEditor
      ref={ref}
      previewFirstPage
      experimentalWorkerOpen={false}
      documentBuffer={
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
      }
      onError={(error) => errors.push(error.message)}
    />
  );
  await waitFor(() => expect(errors).toEqual(['render failed']), { timeout: 10_000 });
  await act(async () => {});
  expect(view.container.querySelector('.docx-editor-error')).not.toBeNull();
  expect(renderer!.displayList).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
  await expectWaitRejects(ref);
}, 30_000);

test('a full session whose canvas replay rejects during preview handover fails the load once', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  holdCanvasReplay = 'full';
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const replayError = new Error('full canvas replay failed');
  const view = render(load(documentBuffer(), (error) => errors.push(error), ref));
  const replay = await currentCanvasReplay();

  expect(created).toBe(2);
  expect(fullSession).not.toBeNull();
  expect(replay.layoutEngine).toBe(fullSession);
  expect(renderer!.presentedEngine).toBe(fullSession);
  expect(renderer!.displayList).toBe(replay.displayList);
  expect(isPresented(renderer!.canvasHostRef.current, replay.displayList)).toBe(false);
  expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
  expect(ref.current!.getDocument()).toBeNull();
  expect(errors).toEqual([]);

  await act(async () => {
    expect(replay.isCurrent()).toBe(true);
    replay.reject(replayError);
  });
  await waitFor(() => expect(errors).toEqual([replayError]), { timeout: 10_000 });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(errors).toEqual([replayError]);
  expect(view.container.querySelector('.docx-editor-error')?.textContent).toContain(
    replayError.message
  );
  expect(view.container.querySelector('.docx-editor-loading')).toBeNull();
  expect(view.queryByTestId('yrs-input')).toBeNull();
  expect(renderer!.displayList).toBeNull();
  expect(renderer!.presentedEngine).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
  await expectWaitRejects(ref);
}, 30_000);

test('a preview whose canvas replay rejects during full-session handover does not fail the load', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  holdCanvasReplay = 'preview';
  let releaseFullDisplayList = () => {};
  holdFullDisplayList = new Promise((done) => (releaseFullDisplayList = done));
  let offReplayFailed = () => {};
  try {
    const ref = createRef<Editor>();
    const errors: Error[] = [];
    const replayError = new Error('preview canvas replay failed');
    const view = render(load(documentBuffer(), (error) => errors.push(error), ref));
    await waitFor(
      () => {
        expect(fullSession).not.toBeNull();
        expect(renderer!.layoutEngine).toBe(fullSession);
      },
      { timeout: 20_000 }
    );
    const replay = await currentCanvasReplay();
    const previewEngine = renderer!.presentedEngine;

    expect(created).toBe(2);
    expect(previewEngine).not.toBeNull();
    expect(previewEngine).not.toBe(fullSession);
    expect((previewEngine as YrsSession).isDisplayOnly()).toBe(true);
    expect(renderer!.layoutEngine).toBe(fullSession);
    expect(renderer!.displayList).toBe(replay.displayList);
    expect(isPresented(renderer!.canvasHostRef.current, replay.displayList)).toBe(false);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(ref.current!.getDocument()).toBeNull();
    expect(errors).toEqual([]);

    const replayFailed = mock((_displayList: object, _error: unknown) => {});
    offReplayFailed = onReplayFailed(replayFailed);
    await act(async () => {
      expect(replay.isCurrent()).toBe(true);
      replay.reject(replayError);
    });
    await waitFor(() => expect(replayFailed).toHaveBeenCalledWith(replay.displayList, replayError));
    await act(async () => {});
    expect(errors).toEqual([]);
    expect(view.container.querySelector('.docx-editor-error')).toBeNull();
    expect(renderer!.layoutEngine).toBe(fullSession);
    expect(renderer!.presentedEngine).toBe(previewEngine);
    expect(renderer!.displayList).toBe(replay.displayList);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(ref.current!.getDocument()).toBeNull();

    holdCanvasReplay = 'full';
    await act(async () => releaseFullDisplayList());
    await waitFor(() => expect(renderer!.presentedEngine).toBe(fullSession), { timeout: 10_000 });
    const fullReplay = await currentCanvasReplay();
    expect(fullReplay.layoutEngine).toBe(fullSession);
    expect(fullReplay.displayList).not.toBe(replay.displayList);
    expect(renderer!.displayList).toBe(fullReplay.displayList);
    expect(isPresented(renderer!.canvasHostRef.current, fullReplay.displayList)).toBe(false);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(ref.current!.getDocument()).toBeNull();

    await act(async () => {
      expect(fullReplay.isCurrent()).toBe(true);
      fullReplay.resolve();
    });
    await waitFor(() =>
      expect(isPresented(renderer!.canvasHostRef.current, fullReplay.displayList)).toBe(true)
    );
    await waitFor(() =>
      expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false)
    );
    expect(errors).toEqual([]);
    expect(view.container.querySelector('.docx-editor-error')).toBeNull();
    expect(renderer!.layoutEngine).toBe(fullSession);
    expect(renderer!.presentedEngine).toBe(fullSession);
    expect(ref.current!.getEditorRef()!.getYrsSession()).toBe(fullSession as YrsSession);
    expect(ref.current!.getDocument()).not.toBeNull();
    expect(ref.current!.getTotalPages()).toBeGreaterThan(0);
    expect(created).toBe(2);
  } finally {
    offReplayFailed();
    releaseFullDisplayList();
    holdFullDisplayList = null;
  }
}, 30_000);

test('an ordinary editor logs a canvas replay rejection without failing the load', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  holdCanvasReplay = 'ordinary';
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const replayError = new Error('ordinary canvas replay failed');
  const previousError = console.error;
  const logged = mock((..._args: unknown[]) => {});
  console.error = logged;
  try {
    const view = render(
      <DocxEditor
        ref={ref}
        previewFirstPage={false}
        experimentalWorkerOpen={false}
        documentBuffer={documentBuffer()}
        onError={(error) => errors.push(error)}
      />
    );
    const replay = await currentCanvasReplay();
    expect(created).toBe(1);
    expect(renderer!.displayList).toBe(replay.displayList);
    expect(isPresented(renderer!.canvasHostRef.current, replay.displayList)).toBe(false);
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false);
    expect(ref.current!.getDocument()).not.toBeNull();

    await act(async () => {
      expect(replay.isCurrent()).toBe(true);
      replay.reject(replayError);
    });
    await waitFor(() =>
      expect(logged).toHaveBeenCalledWith('[CanvasRenderer] Canvas replay failed', replayError)
    );
    await act(async () => {
      await new Promise((done) => setTimeout(done, 200));
    });
    expect(errors).toEqual([]);
    expect(view.container.querySelector('.docx-editor-error')).toBeNull();
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(false);
    expect(renderer!.status).toBe('ready');
    expect(renderer!.displayList).not.toBeNull();
    expect(ref.current!.getTotalPages()).toBeGreaterThan(0);
    expect(ref.current!.getDocument()).not.toBeNull();
  } finally {
    console.error = previousError;
  }
}, 30_000);

test.each(['renderer first', 'replica first'])('renderer and replica notify once per error with %s', async (order) => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const buffer = documentBuffer();
  const onError = (error: Error) => errors.push(error);
  let presented!: () => void;
  const ready = new Promise<void>((resolve) => { presented = resolve; });
  const element = () => (
    <DocxEditor
      ref={ref} previewFirstPage={false} documentBuffer={buffer} onError={onError}
      experimentalWorkerOpen={false}
      onFirstPagePainted={presented}
    />
  );
  const view = render(element());
  await ready;
  await act(async () => {});
  expect(renderer!.status).toBe('ready');
  const session = ref.current!.getEditorRef()!.getYrsSession()!;
  const failure = new Error('resident layout failed');
  const reportReplica = () => replicaError!(failure);
  const reportRenderer = () => {
    workerFailure = { error: failure, errorEngine: session };
    view.rerender(element());
  };
  const reportFirst = order === 'renderer first' ? reportRenderer : reportReplica;
  const reportSecond = order === 'renderer first' ? reportReplica : reportRenderer;
  await act(async () => reportFirst());
  expect(errors).toEqual([failure]);
  renderer!.resetSettled();
  await act(async () => reportSecond());
  expect(errors).toEqual([failure]);
  await act(async () => reportReplica());
  await expect(renderer!.settledDisplayList(null, null)).rejects.toBe(failure);
  expect(errors).toEqual([failure]);
  const next = new Error(failure.message);
  await act(async () => replicaError!(next));
  expect(errors).toEqual([failure, next]);
});

test('a load whose full session fails to lay out fails once, and leaves no session behind', async () => {
  created = 0;
  fullOpen = 'layout-fail';
  failRender = null;
  const ref = createRef<Editor>();
  const errors: string[] = [];
  const view = render(load(documentBuffer(), (error) => errors.push(error.message), ref));
  await waitFor(() => expect(errors).toContain('layout failed'), { timeout: 10_000 });
  await act(async () => {
    await new Promise((done) => setTimeout(done, 200));
  });
  expect(errors).toEqual(['layout failed']);
  expect(view.container.querySelector('.docx-editor-error')).not.toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);
  expect(ref.current!.getDocument()).toBeNull();
}, 30_000);

test.each([false, true])(
  "a preview's render error is reported once and does not fail the full session (workerOpen=%p)",
  async (workerOpen) => {
    created = 0;
    fullOpen = 'open';
    failRender = 'preview';
    const ref = createRef<Editor>();
    const errors: string[] = [];
    const buffer = documentBuffer();
    const view = render(load(buffer, (error) => errors.push(error.message), ref, workerOpen));
    await waitFor(() => expect(created).toBe(2), { timeout: 10_000 });
    await waitFor(() => expect(errors).toEqual(['preview render failed']), { timeout: 10_000 });
    // A host passing a new callback while the full session opens.
    view.rerender(load(buffer, (error) => errors.push(`again: ${error.message}`), ref, workerOpen));
    await waitFor(() => expect(ref.current!.getDocument()).not.toBeNull(), { timeout: 10_000 });
    expect(errors).toEqual(['preview render failed']);
    expect(view.container.querySelector('.docx-editor-error')).toBeNull();
  },
  30_000
);

test('while its preview shows, a load has no page count and its layout is not complete', async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  let release = () => {};
  holdFullOpen = new Promise((done) => (release = done));
  try {
    const ref = createRef<Editor>();
    render(load(documentBuffer(), () => {}, ref));
    await waitFor(() => expect(ref.current).not.toBeNull());
    let pages = null as number | null;
    let totalPages = null as number | null;
    void ref.current!.whenLayoutComplete().then((count) => {
      pages = count;
      totalPages = ref.current!.getTotalPages();
    });
    await waitFor(() => expect(shownPages).toBe(true), { timeout: 10_000 });
    await act(async () => {
      await new Promise((done) => setTimeout(done, 200));
    });
    expect(ref.current!.getTotalPages()).toBe(0);
    expect(pages).toBeNull();

    release();
    await waitFor(() => expect(pages).not.toBeNull(), { timeout: 20_000 });
    expect(renderer!.presentedEngine).toBe(fullSession);
    expect(pages).toBe(renderer!.displayList!.pages.length);
    expect(totalPages).toBe(pages);
  } finally {
    release();
    holdFullOpen = null;
  }
}, 30_000);

test("a preview's layout error is reported and fails no wait for the document", async () => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  failPreviewLayout = true;
  try {
    const ref = createRef<Editor>();
    const errors: string[] = [];
    render(load(documentBuffer(), (error) => errors.push(error.message), ref));
    await waitFor(() => expect(ref.current).not.toBeNull());
    let outcome = null as number | Error | null;
    ref.current!.whenLayoutComplete().then(
      (pages) => (outcome = pages),
      (error: Error) => (outcome = error)
    );
    await waitFor(() => expect(errors).toContain('preview layout failed'), { timeout: 10_000 });
    await waitFor(() => expect(outcome).not.toBeNull(), { timeout: 20_000 });
    expect(renderer!.presentedEngine).toBe(fullSession);
    expect(outcome).toBe(renderer!.displayList!.pages.length);
  } finally {
    failPreviewLayout = false;
  }
}, 30_000);

test.each(['readOnly', 'viewing'] as const)('viewer sessions stay read-only while opening with %s', async (mode) => {
  created = 0;
  fullSession = null;
  shownPages = false;
  fullOpen = 'open';
  failRender = null;
  const startWorker = await residentWorkerFactory();
  const compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(new WebAssembly.Module(
    new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
  ));
  const previousWorker = globalThis.Worker;
  let release = () => {};
  let session: YrsSession | null = null;
  const create = real.createYrsSession;
  const capture = spyOn(real, 'createYrsSession').mockImplementation(async (options) => {
    const pending = await create(options);
    session ??= pending;
    return pending;
  });
  const posted: ResidentEngineWorkerRequest[] = [];
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      const send = worker.postMessage.bind(worker);
      let listener: typeof worker.onmessage = null;
      Object.defineProperty(worker, 'onmessage', {
        get: () => (event: MessageEvent<ResidentEngineWorkerResponse>) => {
          const request = posted.find((request) => request.id === event.data.id);
          if (request?.type === 'bootstrap') release = () => listener?.(event);
          else listener?.(event);
        },
        set: (next: typeof worker.onmessage) => { listener = next; },
      });
      worker.postMessage = (request, transfer) => {
        if (request.type !== 'editModule') posted.push(request);
        send(request, transfer);
      };
      return worker;
    }
  } as unknown as typeof Worker;
  const ref = createRef<Editor>();
  const errors: Error[] = [];
  const view = render(<DocxEditor ref={ref} experimentalWorkerOpen
    readOnly={mode === 'readOnly'} mode={mode === 'viewing' ? 'viewing' : 'editing'}
    documentBuffer={documentBuffer()} onError={(error) => errors.push(error)} />);
  try {
    await waitFor(() => {
      expect(posted.some((request) => request.type === 'bootstrap')).toBe(true);
      expect(session!.isDisplayOnly()).toBe(false);
      expect(renderer!.displayList).toBeNull();
    }, { timeout: 10_000 });
    const pending = session!;
    const version = pending.version();
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    expect(textarea.readOnly).toBe(true);
    expect(() => ref.current!.getDocument()).toThrow(DocxAsyncOnlyError);
    const before = await renderer!.readWorkerDocument({ kind: 'readParagraphs', request: { view: 'accepted' } });
    if (!before.value.ok) throw new Error('The viewer worker could not read its document');
    fireEvent.input(textarea, { target: { value: 'ignored' } });
    fireEvent.keyDown(textarea, { key: 'Enter' });
    fireEvent.keyDown(textarea, { key: 'Backspace' });
    fireEvent.paste(textarea, { clipboardData: { getData: () => 'ignored paste' } });
    await act(async () => {});
    expect(ref.current!.getEditorRef()?.hasPendingInput() ?? false).toBe(false);
    expect(pending.version()).toBe(version);
    await act(async () => release());
    await waitFor(() => {
      expect(renderer!.presentedEngine).toBe(pending);
      expect(isPresented(renderer!.canvasHostRef.current, renderer!.displayList!)).toBe(true);
    }, { timeout: 10_000 });
    expect((view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    const after = await renderer!.readWorkerDocument({ kind: 'readParagraphs', request: { view: 'accepted' } });
    if (!after.value.ok) throw new Error('The viewer worker could not read its document');
    expect(after.value.paragraphs.map((paragraph) => paragraph.text))
      .toEqual(before.value.paragraphs.map((paragraph) => paragraph.text));
    expect(ref.current!.getEditorRef()?.hasPendingInput() ?? false).toBe(false);
    expect(errors).toEqual([]);
    expect(workerOpenDocumentHeld(pending)).toBe(true);
    expect(pending.storyIds()).toEqual([]);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(posted.some((request) => request.type === 'applyInput')).toBe(false);
  } finally {
    view.unmount();
    release();
    capture.mockRestore();
    compileModule.mockRestore();
    globalThis.Worker = previousWorker;
  }
}, 30_000);
