import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { useCallback, useEffect, useRef, useState } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsSession,
  preloadResidentEngineWorker,
  ResidentWorkerOutOfMemoryError,
  type YrsDocxHost,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import type { ResidentEngineWorkerRequest } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { LayoutSelectionGate } from '@betteroffice/docx/layout';
import { useCanvasRenderer, type OpenInWorker } from './useDisplayList';
import { useLayoutPipeline } from './useLayoutPipeline';
import type { DocxHostSearch } from './useHostSearch';
import { useYrsCoreSession } from './useYrsCoreSession';
import type { DocxEditorCollaborationOptions } from '../types';
import { awaitWorkerOpenReplica, ensureWorkerOpenReplica } from '../internals/workerOpenReplica';
import { sourceVersionOf } from '../internals/layoutProvenance';
import * as replicaHelpers from '../internals/workerOpenReplica';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import { createCommentIdAllocator } from '../commentFactories';
import { useDocxEditorRefApi } from './useDocxEditorRefApi';
import { usePagedEditorCommandBridge, type PagedEditorCommandBridge } from './usePagedEditorRefApi';
import { YrsInput, type YrsInputRef } from '../YrsInput';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;
const bytes = new Uint8Array(readFileSync(resolve(
  import.meta.dir,
  '../../../../../../crates/docx-edit/tests/fixtures/page-fragments/pages.docx'
)));

async function longFixture(paragraphs = 205): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  );
  const body = Array.from({ length: paragraphs }, (_, index) => {
    const text = index === 0 ? 'First paragraph' : index === paragraphs - 1 ? 'Tail paragraph' : `Paragraph ${index}`;
    return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  }).join('');
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`
  );
  return zip.generateAsync({ type: 'uint8array' });
}

const longBytes = await longFixture();
const font = new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));
const sessions: YrsSession[] = [];
let startWorker!: () => InProcessResidentWorker;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});
afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function installWorker(options: {
  failOpen?: boolean;
  failState?: boolean;
  holdState?: boolean;
  holdOpen?: boolean;
  oomStage?: 'open' | 'fontRequirements' | 'bootstrap' | 'encodeState';
  holdRetryOpen?: boolean;
  revisionCount?: number;
  failRevisionCount?: boolean;
  onRevisionCount?: () => void;
  holdCompletion?: boolean;
} = {}) {
  const workers: InProcessResidentWorker[] = [];
  const posted: ResidentEngineWorkerRequest[] = [];
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      const send = worker.postMessage.bind(worker);
      worker.postMessage = (request, transfer) => {
        posted.push(request);
        if ((options.holdState && request.type === 'encodeState') ||
            (options.holdOpen && request.type === 'open') ||
            (options.holdRetryOpen && workers.length > 1 && request.type === 'open') ||
            (options.holdCompletion && request.type === 'completeLayout')) worker.hold();
        if (options.oomStage === request.type &&
            (request.type !== 'encodeState' || workers.length === 1)) {
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'worker exhausted memory', terminal: true, outOfMemory: true },
          } as MessageEvent));
        } else if ((options.failOpen && request.type === 'open') ||
            (options.failState && request.type === 'encodeState')) {
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'open failed', terminal: true },
          } as MessageEvent));
        } else if (request.type === 'revisionCount') {
          options.onRevisionCount?.();
          queueMicrotask(() => worker.onmessage?.({
            data: options.failRevisionCount
              ? { id: request.id, ok: false, error: 'revision count failed' }
              : { id: request.id, ok: true, revisionCount: options.revisionCount ?? 0 },
          } as MessageEvent));
        } else send(request, transfer);
      };
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  return { workers, posted };
}

interface HarnessProps {
  experimentalWorkerOpen: boolean;
  hydrateOnDemand?: boolean;
  previewFirstPage?: boolean;
  openInWorker?: OpenInWorker;
  source: Uint8Array;
  generation: number;
  collaboration?: DocxEditorCollaborationOptions;
  readOnly?: boolean;
  resolvedCommentIds?: ReadonlySet<number>;
  /** Asks for the replica as soon as the session exists, as DocxEditor does for plugins, sidebars or the outline. */
  wanted?: boolean;
  /** Passes the renderer's own pending completion, as DocxEditor does. */
  followCompletion?: boolean;
  /** Holds the replica as while the shown engine's completion is still to be asked of the worker. */
  holdReplica?: boolean;
}

function useHarness(props: HarnessProps) {
  const relayout = useRef<(() => void) | null>(null);
  const handoffFromRef = useRef<YrsSession | null>(null);
  const renderer = useCanvasRenderer(
    undefined,
    props.resolvedCommentIds,
    () => relayout.current?.(),
    undefined,
    handoffFromRef,
    props.experimentalWorkerOpen
  );
  useEffect(() => renderer.resetSettled(), [props.generation]);
  const [host, setHost] = useState<YrsDocxHost | null>(null);
  const mainOpens = useRef<boolean[]>([]);
  const errors = useRef<Error[]>([]);
  const loadChecks = useRef<number[]>([]);
  const openInWorker = useCallback<OpenInWorker>((session, source, digest, generation) => {
    return (props.openInWorker ?? renderer.openInWorker)(session, source, digest, generation);
  }, [props.openInWorker, renderer.openInWorker]);
  const core = useYrsCoreSession(
    true, host?.document ?? null, null, props.source, props.generation, props.collaboration,
    {
      isCurrentLoad: (generation) => {
        loadChecks.current.push(generation);
        return generation === props.generation;
      },
      onSession: (session) => {
        renderer.recordSession(session);
        const open = session.openDocx.bind(session);
        session.openDocx = (input, seed, options) => {
          mainOpens.current.push(seed);
          return open(input, seed, options);
        };
      },
      onHostDocument: setHost,
      onError: (error) => errors.current.push(error),
    },
    {
      previewFirstPage: props.previewFirstPage,
      heldEngine: renderer.layoutEngine,
      shownEngine: renderer.presentedEngine,
      workerOpen: props.experimentalWorkerOpen ? {
        openInWorker,
        renderedFrame: renderer.status === 'ready' ? renderer.displayList : null,
        ...(props.holdReplica ? { pendingCompletion: renderer.presentedEngine } : {}),
        ...(props.followCompletion ? { pendingCompletion: renderer.pendingCompletion } : {}),
        hydrateOnDemand: props.hydrateOnDemand,
      } : undefined,
    }
  );
  handoffFromRef.current = core.handoffFrom;
  const replicaPending = Boolean(core.hydrateOnDemand && core.session && !core.replicaReady);
  useEffect(() => {
    if (props.wanted && replicaPending) core.requestReplica();
  }, [core, props.wanted, replicaPending]);
  const syncCoordinator = useRef(new LayoutSelectionGate());
  const element = useRef<HTMLDivElement | null>(null);
  const registeredFont = useRef<{ session: YrsSession; id: number } | null>(null);
  const pipeline = useLayoutPipeline({
    document: host?.document ?? null,
    session: core.session,
    renderEnv: {},
    pageGap: 24,
    zoom: 1,
    residentMeasurementConfig: (requirements) => {
      const session = core.session;
      if (!session) return null;
      if (registeredFont.current?.session !== session) {
        registeredFont.current = { session, id: session.registerFont(font) };
      }
      const id = registeredFont.current.id;
      return {
        fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [id]])),
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: { noLeading: false, doNotExpandShiftReturn: false },
        authoritativeShaping: true,
      };
    },
    deferLayoutPass: () => false,
    pagesContainerRef: element,
    viewportLayoutRef: element,
    syncCoordinator: syncCoordinator.current,
    getScrollContainer: () => null,
    onLayoutComputed: (layout) => renderer.onLayoutComputed(layout, core.session),
    layoutInWorker: renderer.layoutInWorker,
    experimentalWorkerOpen: props.experimentalWorkerOpen,
    fontRequirementsInWorker: props.experimentalWorkerOpen ? renderer.fontRequirementsInWorker : undefined,
    onError: (error, session) => {
      if (!core.failOpening(error, session)) errors.current.push(error);
    },
  });
  relayout.current = pipeline.runLayoutPipeline;
  useEffect(() => {
    if (props.readOnly && core.session) relayout.current?.();
  }, [props.readOnly, core.session]);
  const pagedEditorRef = useRef<PagedEditorRef | null>(null);
  pagedEditorRef.current = core.session ? {
    getYrsSession: () => core.session,
    getDocument: core.documentFromYrs,
    flushPendingInput: async () => {},
  } as PagedEditorRef : null;
  const ref = useRef<DocxEditorRef>(null);
  useDocxEditorRefApi({
    experimentalWorkerOpen: props.experimentalWorkerOpen,
    ref,
    document: host?.document ?? null,
    documentFromYrs: core.documentFromYrs,
    historyStateRef: { current: host?.document ?? null },
    pagedEditorRef,
    handleSave: async () => core.documentFromYrs() ? new ArrayBuffer(0) : null,
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
    getCachedStyleResolver: (() => { throw new Error('unused'); }) as never,
    commentIdAllocator: createCommentIdAllocator(),
    commands: UNAVAILABLE_DOCX_COMMANDS,
    modeRef: { current: 'viewing' },
    allowHostProposalsRef: { current: false },
    hostSearch: {} as DocxHostSearch,
  });
  const bridgeRef = useRef<PagedEditorCommandBridge | null>(null);
  const inputRef = useRef<YrsInputRef | null>(null);
  inputRef.current = {
    runAfterPendingInput: async (operation) => operation(),
    hasPendingInput: () => false,
  } as YrsInputRef;
  const selectionRef = useRef(null);
  const listenersRef = useRef(new Set<() => void>());
  usePagedEditorCommandBridge({
    experimentalWorkerOpen: props.experimentalWorkerOpen,
    bridgeRef,
    yrsInputRef: inputRef,
    session: core.session,
    rootStory: 'body',
    inputPositionMap: core.inputPositionMap,
    latestSelectionRef: selectionRef,
    listenersRef,
    getPositionProjection: () => null,
    displayPositionToLoc: () => null,
    format: () => false,
    command: () => false,
    syncYrsInputState: () => false,
    yrsLocToDisplayPosition: () => null,
    scrollToPositionImpl: () => {},
  });
  return {
    core, renderer, pipeline, host, ref, bridgeRef,
    presentFrame: () => core.notifyFramePresented(renderer.presentedEngine),
    loadChecks: loadChecks.current, mainOpens: mainOpens.current, errors: errors.current,
  };
}

const initialProps: HarnessProps = { experimentalWorkerOpen: true, source: bytes, generation: 1 };

function texts(session: YrsSession) {
  return Object.fromEntries(session.storyIds().sort().map((story) => [
    story, session.paragraphs(story).map((paragraph) => paragraph.text),
  ]));
}

function holdFrames() {
  const request = globalThis.requestAnimationFrame;
  const cancel = globalThis.cancelAnimationFrame;
  const frames = new Map<number, FrameRequestCallback>();
  let nextId = 0;
  globalThis.requestAnimationFrame = (callback) => {
    frames.set(++nextId, callback);
    return nextId;
  };
  globalThis.cancelAnimationFrame = (id) => { frames.delete(id); };
  return {
    run() {
      const pending = [...frames.values()];
      frames.clear();
      for (const callback of pending) callback(performance.now());
    },
    restore() {
      globalThis.requestAnimationFrame = request;
      globalThis.cancelAnimationFrame = cancel;
    },
  };
}

test.each([false, true])('textarea focus requests a replica only with hydrateOnDemand=%s', async (hydrateOnDemand) => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, hydrateOnDemand },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const core = result.current.core;
    expect(core.hydrateOnDemand).toBe(hydrateOnDemand);
    if (hydrateOnDemand) {
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
      act(() => result.current.presentFrame());
      act(() => frames.run());
      act(() => frames.run());
      await waitFor(() => expect(posted.map((request) => request.type)).toContain('revisionCount'));
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    }
    const view = render(
      <YrsInput
        enabled
        readOnly={hydrateOnDemand}
        session={core.session}
        replicaReadyRef={core.replicaReadyRef}
        requestReplica={core.hydrateOnDemand ? core.requestReplica : undefined}
        inputPositionMap={core.inputPositionMap}
        displayPositionToLoc={() => null}
        locToDisplayPosition={() => null}
        onStateChange={() => {}}
        onDirectInput={() => {}}
      />
    );
    const textarea = view.getByTestId('yrs-input');
    act(() => {
      requestAnimationFrame(() => textarea.focus());
      frames.run();
      fireEvent.keyDown(textarea, { key: 'ArrowRight' });
    });
    await act(async () => { await Promise.resolve(); });
    if (hydrateOnDemand) {
      await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
      await act(async () => {
        workers[0].release();
        await awaitWorkerOpenReplica(core.session!);
      });
      await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    } else {
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
      expect(result.current.core.replicaReady).toBe(false);
    }
  } finally {
    unmount();
    cleanup();
    frames.restore();
  }
});

test('on-demand hydration is inactive without worker opening', async () => {
  installWorker();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, experimentalWorkerOpen: false, hydrateOnDemand: true },
  });
  try {
    await waitFor(() => expect(result.current.core.session).not.toBeNull());
    expect(result.current.core.hydrateOnDemand).toBe(false);
  } finally {
    unmount();
  }
});

test('an on-demand replica stays empty past its load point until requested', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames();
  try {
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, hydrateOnDemand: true },
    });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1));
    await act(async () => {});
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.core.replicaReadyRef?.current).toBe(false);
    expect(session.storyIds()).toEqual([]);
    expect(result.current.mainOpens).toEqual([]);
    act(() => result.current.core.requestReplica());
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.core.replicaReadyRef?.current).toBe(true);
    expect(session.hasStory('body')).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
}, 15_000);

test('tracked changes start an on-demand replica without a replica request', async () => {
  const { workers, posted } = installWorker({ holdState: true, revisionCount: 1 });
  const frames = holdFrames();
  const props = { ...initialProps, hydrateOnDemand: true, holdReplica: true };
  const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await act(async () => {});
    expect(posted.some((request) => request.type === 'revisionCount')).toBe(false);
    rerender({ ...props, holdReplica: false });
    expect(posted.some((request) => request.type === 'revisionCount')).toBe(false);
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
    rerender({ ...props, holdReplica: true });
    rerender({ ...props, holdReplica: false });
    act(() => frames.run());
    act(() => frames.run());
    await act(async () => {});
    expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
    expect(result.current.core.replicaReady).toBe(false);
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(session.hasStory('body')).toBe(true);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    cleanup();
    frames.restore();
  }
});

test.each(['wanted', 'awaited'] as const)(
  'a replica %s before the first layout loads only once the rest of the layout is asked of the worker',
  async (how) => {
    const { workers, posted } = installWorker({ holdState: true, holdCompletion: true });
    const frames = holdFrames();
    const { result, unmount } = renderHook(useHarness, {
      initialProps: {
        ...initialProps,
        source: await longFixture(1200),
        hydrateOnDemand: true,
        followCompletion: true,
        wanted: how === 'wanted',
      },
    });
    try {
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const session = result.current.core.session!;
      const ready = how === 'awaited' ? awaitWorkerOpenReplica(session) : undefined;
      await act(async () => {});
      expect(posted.map((request) => request.type)).toEqual(['open']);
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(posted.map((request) => request.type)).toContain('completeLayout'), {
        timeout: 5000,
      });
      await waitFor(() => expect(result.current.renderer.pendingCompletion).toBeNull());
      await act(async () => {});
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
      act(() => result.current.presentFrame());
      act(() => frames.run());
      act(() => frames.run());
      await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
      const types = posted.map((request) => request.type);
      expect(types.indexOf('encodeState')).toBeGreaterThan(types.indexOf('bootstrap'));
      expect(types.indexOf('encodeState')).toBeGreaterThan(types.indexOf('completeLayout'));
      await act(async () => {
        workers[0].release();
        await (ready ?? awaitWorkerOpenReplica(session));
      });
      await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
      expect(result.current.mainOpens).toEqual([false]);
      expect(result.current.errors).toEqual([]);
    } finally {
      unmount();
      cleanup();
      frames.restore();
    }
  },
  15_000
);

test('the revision count is asked once the rest of the layout is asked of the worker, not after it completes', async () => {
  const { workers, posted } = installWorker({ holdCompletion: true });
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps,
      source: await longFixture(1200),
      hydrateOnDemand: true,
      followCompletion: true,
    },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(posted.map((request) => request.type)).toContain('completeLayout'), {
      timeout: 5000,
    });
    await waitFor(() => expect(result.current.renderer.pendingCompletion).toBeNull());
    await act(async () => {});
    expect(posted.some((request) => request.type === 'revisionCount')).toBe(false);
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1));
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    await act(async () => {
      workers[0].release();
    });
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await act(async () => {});
    expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    cleanup();
    frames.restore();
  }
}, 15_000);

test('a failed revision count starts the on-demand replica', async () => {
  const { workers, posted } = installWorker({ holdState: true, failRevisionCount: true });
  const frames = holdFrames();
  const props = { ...initialProps, hydrateOnDemand: true, holdReplica: true };
  const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await act(async () => {});
    expect(posted.some((request) => request.type === 'revisionCount')).toBe(false);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    rerender({ ...props, holdReplica: false });
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
    expect(result.current.core.replicaReady).toBe(false);
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(session.hasStory('body')).toBe(true);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    cleanup();
    frames.restore();
  }
});

test('turning off on-demand hydration starts a pending replica after two frames', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames();
  try {
    const props = { ...initialProps, hydrateOnDemand: true };
    const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const requestReplica = result.current.core.requestReplica;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => rerender({ ...props, hydrateOnDemand: false }));
    expect(result.current.core.session).toBe(session);
    expect(result.current.core.requestReplica).toBe(requestReplica);
    expect(result.current.core.hydrateOnDemand).toBe(false);
    expect(replicaHelpers.workerOpenReplicaOnDemand(session)).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    act(() => frames.run());
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => frames.run());
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test('a painted main preview hands off to the worker before hydrating the full replica', async () => {
  const { workers, posted } = installWorker({ holdOpen: true, holdState: true });
  const frames = holdFrames();
  try {
    const props = { ...initialProps, previewFirstPage: true, source: longBytes };
    const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    const destroyed = spyOn(preview, 'destroy');
    expect(preview.isDisplayOnly()).toBe(true);
    expect(preview.paragraphs('body').map((paragraph) => paragraph.text)).not.toContain('Tail paragraph');
    expect(result.current.core.replicaReady).toBe(true);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    expect(result.current.renderer.status).toBe('ready');
    // The worker opens the full document while the preview opens and paints.
    await waitFor(() => expect(posted.map((request) => request.type)).toEqual(['open']));
    expect(workers).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([]);

    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    expect(posted.map((request) => request.type)).toEqual(['open']);
    expect(result.current.core.session).toBe(preview);
    expect(destroyed).not.toHaveBeenCalled();
    const firstPreviewFrame = result.current.renderer.displayList;
    rerender({ ...props, resolvedCommentIds: new Set([1]) });
    await waitFor(() => expect(result.current.renderer.displayList).not.toBe(firstPreviewFrame));
    expect(posted.map((request) => request.type)).toEqual(['open']);
    await act(async () => { workers[0].release(); });
    await waitFor(() => expect(result.current.core.previewing).toBe(false));
    const full = result.current.core.session!;
    expect(full).not.toBe(preview);
    expect(full.isDisplayOnly()).toBe(false);
    expect(full.storyIds()).toEqual([]);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.core.replicaReadyRef?.current).toBe(false);
    expect(result.current.core.handoffFrom).toBe(preview);
    expect(result.current.core.opening).toBe(true);
    act(() => frames.run());
    act(() => frames.run());
    expect(posted.map((request) => request.type)).toEqual(['open']);

    const inheritedPreviewFrame = result.current.renderer.displayList;
    rerender({ ...props, resolvedCommentIds: new Set([2]) });
    await waitFor(() => expect(result.current.renderer.displayList).not.toBe(inheritedPreviewFrame));
    expect(result.current.renderer.presentedEngine).toBe(preview);
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    expect(posted.map((request) => request.type)).toEqual(['open']);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.core.handoffFrom).toBe(preview);
    expect(destroyed).not.toHaveBeenCalled();

    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
    expect(posted.map((request) => request.type).slice(0, 3)).toEqual(['open', 'fontRequirements', 'bootstrap']);
    expect(posted.find((request) => request.type === 'bootstrap')).toMatchObject({ opened: true });
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.core.handoffFrom).toBe(preview);
    expect(destroyed).not.toHaveBeenCalled();
    act(() => result.current.presentFrame());
    expect(result.current.core.handoffFrom).toBeNull();
    expect(result.current.core.opening).toBe(false);
    expect(destroyed).toHaveBeenCalledTimes(1);
    act(() => frames.run());
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => frames.run());
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.core.replicaReady).toBe(false);
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(full);
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.core.replicaReadyRef?.current).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    const paragraphs = full.paragraphs('body');
    expect(paragraphs).toHaveLength(205);
    expect(paragraphs[0].text).toBe('First paragraph');
    expect(paragraphs.at(-1)!.text).toBe('Tail paragraph');
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test.each(['null', 'throw'] as const)(
  'a combined worker open keeps the preview through main fallback with outcome=%s',
  async (outcome) => {
    installWorker();
    const frames = holdFrames();
    let release!: () => void;
    const held = new Promise<void>((resolve) => { release = resolve; });
    let calls = 0;
    const openInWorker: OpenInWorker = async () => {
      calls += 1;
      await held;
      if (outcome === 'throw') throw new Error('worker unavailable');
      return null;
    };
    try {
      const { result, unmount } = renderHook(useHarness, {
        initialProps: { ...initialProps, previewFirstPage: true, source: longBytes, openInWorker },
      });
      await waitFor(() => expect(result.current.core.previewing).toBe(true));
      const preview = result.current.core.session!;
      const destroyed = spyOn(preview, 'destroy');
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
      await waitFor(() => expect(calls).toBe(1));
      expect(result.current.core.session).toBe(preview);
      act(() => result.current.presentFrame());
      act(() => frames.run());
      act(() => frames.run());
      await waitFor(() => expect(calls).toBe(1));
      expect(result.current.core.session).toBe(preview);
      expect(result.current.renderer.presentedEngine).toBe(preview);
      expect(result.current.mainOpens).toEqual([]);
      expect(destroyed).not.toHaveBeenCalled();
      await act(async () => { release(); });
      await waitFor(() => expect(result.current.core.previewing).toBe(false));
      const full = result.current.core.session!;
      expect(full).not.toBe(preview);
      expect(result.current.mainOpens).toEqual([true]);
      expect(result.current.core.replicaReady).toBe(true);
      expect(result.current.core.handoffFrom).toBe(preview);
      expect(result.current.renderer.presentedEngine).toBe(preview);
      expect(destroyed).not.toHaveBeenCalled();
      act(() => result.current.presentFrame());
      expect(result.current.core.handoffFrom).toBe(preview);
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
      expect(destroyed).not.toHaveBeenCalled();
      act(() => result.current.presentFrame());
      expect(result.current.core.handoffFrom).toBeNull();
      expect(destroyed).toHaveBeenCalledTimes(1);
      expect(result.current.errors).toEqual([]);
      unmount();
    } finally {
      cleanup();
      frames.restore();
    }
  }
);

test('terminal worker open OOM fails a previewing load once without a main seed', async () => {
  const { workers, posted } = installWorker({ oomStage: 'open' });
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, previewFirstPage: true, source: longBytes },
  });
  await waitFor(() => expect(result.current.core.previewing).toBe(true));
  act(() => result.current.pipeline.runLayoutPipeline());
  await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(result.current.core.session));
  act(() => result.current.presentFrame());
  await waitFor(() => expect(result.current.errors).toHaveLength(1));
  expect(result.current.errors[0]).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
  expect(workers).toHaveLength(2);
  expect(posted.map((request) => request.type)).toEqual(['open', 'open']);
  expect(result.current.mainOpens).toEqual([]);
  expect(result.current.core.session).toBeNull();
  expect(result.current.core.opening).toBe(false);
  expect(result.current.core.handoffFrom).toBeNull();
  expect(result.current.core.failOpening(new Error('later'))).toBe(false);
  expect(result.current.errors).toHaveLength(1);
});

test('failure of the accepted full session before its frame cancels deferred hydration', async () => {
  const { workers, posted } = installWorker();
  const frames = holdFrames();
  try {
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, previewFirstPage: true, source: longBytes },
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(result.current.core.previewing).toBe(false));
    const full = result.current.core.session!;
    expect(replicaHelpers.workerOpenReplicaPending(full)).toBe(true);
    expect(result.current.core.handoffFrom).toBe(preview);
    expect(result.current.renderer.presentedEngine).toBe(preview);
    const destroyed = spyOn(workers[0], 'terminate');
    const pending = awaitWorkerOpenReplica(full)!;
    const failure = new Error('full frame failed');
    expect(result.current.core.failOpening(new Error('late preview error'), preview)).toBe(false);
    act(() => { expect(result.current.core.failOpening(failure, full)).toBe(true); });
    const rejection = await pending.then(() => null, (error: unknown) => error);
    expect(rejection).toMatchObject({ message: expect.stringContaining('document changed') });
    expect(replicaHelpers.workerOpenReplicaPending(full)).toBe(false);
    expect(destroyed).toHaveBeenCalledTimes(1);
    expect(result.current.core.session).toBeNull();
    expect(result.current.core.handoffFrom).toBeNull();
    expect(result.current.core.opening).toBe(false);
    act(() => frames.run());
    act(() => frames.run());
    expect(posted.map((request) => request.type)).toEqual(['open', 'destroy']);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.core.replicaReadyRef?.current).toBe(false);
    expect(result.current.errors).toEqual([failure]);
    expect(result.current.core.failOpening(new Error('later'), full)).toBe(false);
    expect(result.current.errors).toHaveLength(1);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test('a preloaded spare worker takes the open that starts alongside the preview', async () => {
  const { workers, posted } = installWorker();
  await preloadResidentEngineWorker();
  expect(workers).toHaveLength(1);
  const frames = holdFrames();
  try {
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, previewFirstPage: true, source: longBytes },
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    await waitFor(() => expect(posted.some((request) => request.type === 'open')).toBe(true));
    const preview = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(result.current.core.previewing).toBe(false));
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(result.current.core.session));
    expect(workers).toHaveLength(1);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test('the replica waits while the shown engine is still to ask the worker for the rest of its layout', async () => {
  const { posted } = installWorker();
  const frames = holdFrames();
  try {
    const props = { ...initialProps, source: longBytes, holdReplica: true };
    const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const full = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await act(async () => {});
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    rerender({ ...props, holdReplica: false });
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(posted.some((request) => request.type === 'encodeState')).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test('the default open calls no worker open, font preflight or state handoff', async () => {
  const { workers, posted } = installWorker();
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, experimentalWorkerOpen: false },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.core.replicaReady).toBe(true);
  expect(workers).toHaveLength(0);
  expect(posted).toHaveLength(0);
});

test.each([false, true])(
  'worker font preflight and the first frame precede the main replica with previewFirstPage=%s and collaboration',
  async (previewFirstPage) => {
    const { workers, posted } = installWorker({ holdState: true });
    const replicas: Array<YrsSession | null> = [];
    const { result } = renderHook(useHarness, {
      initialProps: {
        ...initialProps,
        previewFirstPage,
        collaboration: { clientId: 9401, onReplica: (replica) => replicas.push(replica as YrsSession | null) },
      },
    });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    expect(session.clientId).toBe(9401);
    expect(session.storyIds()).toEqual([]);
    expect(result.current.mainOpens).toEqual([]);
    expect(replicas).toEqual([]);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    expect(result.current.errors).toEqual([]);
    expect(posted.map((request) => request.type).slice(0, 3)).toEqual(['open', 'fontRequirements', 'bootstrap']);
    expect(posted.find((request) => request.type === 'bootstrap')).toMatchObject({ opened: true });
    expect(result.current.renderer.frame).not.toBeNull();
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.core.replicaReadyRef?.current).toBe(false);
    expect(session.storyIds()).toEqual([]);

    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.core.replicaReadyRef?.current).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(replicas).toEqual([session]);
    await waitFor(() => expect(sourceVersionOf(result.current.renderer.queries)).toBe(session.version()));
    const direct = await createYrsSession();
    sessions.push(direct);
    direct.openDocx(bytes, true);
    expect(texts(session)).toEqual(texts(direct));
    expect(result.current.core.documentFromYrs()).not.toBeNull();
  }
);

test('a failed worker open falls back to the existing main open', async () => {
  const { posted } = installWorker({ failOpen: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(result.current.mainOpens).toEqual([true]);
  expect(result.current.core.replicaReady).toBe(true);
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.errors).toEqual([]);
  expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
});

test('a comment visibility change during the worker open keeps the opening worker', async () => {
  const { workers, posted } = installWorker({ holdOpen: true });
  const { result, rerender } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(posted.some((request) => request.type === 'open')).toBe(true));
  rerender({ ...initialProps, resolvedCommentIds: new Set([1]) });
  await act(async () => {
    workers[0].release();
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(workers).toHaveLength(1);
  expect(result.current.mainOpens).toEqual([]);
  expect(result.current.errors).toEqual([]);
});

test('an unavailable worker falls back before publishing host metadata', async () => {
  globalThis.Worker = undefined as unknown as typeof Worker;
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(result.current.mainOpens).toEqual([true]);
  expect(result.current.core.replicaReady).toBe(true);
  expect(result.current.errors).toEqual([]);
});

test('a worker lost during the handoff opens a full main replica', async () => {
  installWorker({ failState: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  act(() => result.current.pipeline.runLayoutPipeline());
  await waitFor(() => expect(result.current.mainOpens).toEqual([true]));
  await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.errors).toEqual([]);
});

test('a first-layout font setup failure starts the replica for pending reads, save and commands', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const session = result.current.core.session!;
  const failure = new Error('font setup failed');
  const registerFont = spyOn(session, 'registerFont').mockImplementation(() => { throw failure; });
  try {
    const calls = Promise.allSettled([
      result.current.ref.current!.readParagraphs({ view: 'accepted' }),
      result.current.ref.current!.save(),
      result.current.bridgeRef.current!.runAfterPendingInput(() => true),
    ]);
    const completed = { value: false };
    void calls.then(() => { completed.value = true; });
    expect(result.current.core.failOpening(failure, {})).toBe(false);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true), { timeout: 500 });
    expect(result.current.errors).toEqual([failure]);
    expect(result.current.renderer.frame).toBeNull();
    expect(result.current.mainOpens).toEqual([]);
    expect(completed.value).toBe(false);
    act(() => { result.current.core.failOpening(failure, session); });
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    await act(async () => { workers[0].release(); });
    await waitFor(() => expect(completed.value).toBe(true));
    const settled = await calls;
    expect(settled[0]).toMatchObject({ status: 'fulfilled', value: { ok: true } });
    expect(settled[1]).toEqual({ status: 'fulfilled', value: new ArrayBuffer(0) });
    expect(settled[2]).toEqual({ status: 'fulfilled', value: true });
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.renderer.frame).toBeNull();
  } finally {
    registerFont.mockRestore();
  }
});

test('an on-demand replica requested before any frame loads after the bounded wait', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, hydrateOnDemand: true, wanted: true },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    await act(async () => {});
    expect(posted.map((request) => request.type)).toEqual(['open']);
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true), {
      timeout: 7000,
    });
    const types = posted.map((request) => request.type);
    expect(types.indexOf('revisionCount')).toBeGreaterThan(-1);
    expect(types.indexOf('revisionCount')).toBeLessThan(types.indexOf('encodeState'));
    expect(result.current.renderer.frame).toBeNull();
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    cleanup();
  }
}, 15_000);

test('a worker open without a frame or error starts the replica after the bounded wait', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const calls = Promise.allSettled([
    result.current.ref.current!.readParagraphs({ view: 'accepted' }),
    result.current.ref.current!.save(),
  ]);
  const completed = { value: false };
  void calls.then(() => { completed.value = true; });
  expect(posted.map((request) => request.type)).toEqual(['open']);
  await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true), { timeout: 7000 });
  expect(result.current.renderer.frame).toBeNull();
  expect(result.current.errors).toEqual([]);
  expect(result.current.mainOpens).toEqual([]);
  expect(completed.value).toBe(false);
  await act(async () => { workers[0].release(); });
  await waitFor(() => expect(completed.value).toBe(true));
  const settled = await calls;
  expect(settled[0]).toMatchObject({ status: 'fulfilled', value: { ok: true } });
  expect(settled[1]).toEqual({ status: 'fulfilled', value: new ArrayBuffer(0) });
  expect(result.current.mainOpens).toEqual([false]);
  expect(result.current.core.replicaReady).toBe(true);
}, 15_000);

test('a failed fallback reports the same document error as a normal open', async () => {
  const invalid = Uint8Array.of(1, 2, 3);
  const direct = await createYrsSession();
  sessions.push(direct);
  let expected = '';
  try { direct.openDocx(invalid, true); }
  catch (error) { expected = error instanceof Error ? error.message : String(error); }
  expect(expected).not.toBe('');
  installWorker({ failOpen: true });
  const { result } = renderHook(useHarness, { initialProps: { ...initialProps, source: invalid } });
  await waitFor(() => expect(result.current.errors).toHaveLength(1));
  expect(result.current.errors[0].message).toBe(expected);
  expect(result.current.core.session).toBeNull();
});

test('a shared collaboration update keeps the existing join path', async () => {
  const shared = await createYrsSession();
  sessions.push(shared);
  shared.openDocx(bytes, true);
  const { workers } = installWorker();
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, collaboration: { initialUpdate: shared.encodeState() } },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(workers).toHaveLength(0);
  expect(result.current.core.replicaReady).toBe(true);
  expect(texts(result.current.core.session!)).toEqual(texts(shared));
});

test('main-thread layout fallback opens the pending replica before measuring it', async () => {
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  act(() => result.current.pipeline.runLayoutPipeline({ onHost: true }));
  await waitFor(() => expect(result.current.pipeline.layout).not.toBeNull());
  expect(result.current.mainOpens).toEqual([true]);
  expect(result.current.core.replicaReady).toBe(true);
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.errors).toEqual([]);
});

test('a replaced worker open never publishes its host or revives its replica', async () => {
  const { workers, posted } = installWorker({ holdOpen: true });
  const { result, rerender } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(workers).toHaveLength(1));
  act(() => rerender({ ...initialProps, generation: 2 }));
  await waitFor(() => expect(workers).toHaveLength(2));
  await act(async () => {
    workers[0].release();
    workers[1].release();
  });
  await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
  expect(posted.filter((request) => request.type === 'open').map((request) => request.generation)).toEqual(['1', '2']);
  expect(result.current.mainOpens).toEqual([]);
  expect(result.current.errors).toEqual([]);
  act(() => ensureWorkerOpenReplica(result.current.core.session!));
});

for (const stage of ['fontRequirements', 'bootstrap'] as const) {
  test(`terminal OOM during ${stage} rejects pending reads, save and commands`, async () => {
    const { workers, posted } = installWorker({ oomStage: stage });
    const { result } = renderHook(useHarness, { initialProps });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const calls = Promise.allSettled([
      result.current.ref.current!.readParagraphs({ view: 'accepted' }),
      result.current.ref.current!.save(),
      result.current.ref.current!.flushPendingInput(),
      result.current.bridgeRef.current!.runAfterPendingInput(() => true),
    ]);
    const completed = { value: false };
    void calls.then(() => { completed.value = true; });
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(completed.value).toBe(true));
    await waitFor(() => expect(result.current.renderer.error).toBeInstanceOf(ResidentWorkerOutOfMemoryError));
    const failure = result.current.renderer.error;
    expect(failure).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    const settled: PromiseSettledResult<unknown>[] = await calls;
    expect(settled).toEqual(Array.from({ length: 4 }, () => ({ status: 'rejected', reason: failure })));
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(false);
    expect(result.current.renderer.workerMemory()).toBeNull();
    expect(workers).toHaveLength(2);
    const count = posted.length;
    await expect(result.current.renderer.openInWorker(session, bytes)).rejects.toBe(failure);
    expect(posted).toHaveLength(count);
    expect(result.current.mainOpens).toEqual([]);
  });

  for (const action of ['reload', 'unmount'] as const) {
    test(`${action} during ${stage} OOM retry rejects pending replica calls`, async () => {
      const { workers, posted } = installWorker({ oomStage: stage, holdRetryOpen: true });
      const { result, rerender, unmount } = renderHook(useHarness, { initialProps });
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const session = result.current.core.session!;
      const pending = Promise.allSettled([
        result.current.ref.current!.readParagraphs({ view: 'accepted' }),
        result.current.ref.current!.save(),
        result.current.bridgeRef.current!.runAfterPendingInput(() => true),
      ]);
      const completed = { value: false };
      void pending.then(() => { completed.value = true; });
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(workers).toHaveLength(2));
      await waitFor(() => expect(posted.filter((request) => request.type === 'open')).toHaveLength(2));
      expect(completed.value).toBe(false);
      act(() => {
        if (action === 'reload') rerender({ ...initialProps, generation: 2 });
        else unmount();
      });
      await waitFor(() => expect(completed.value).toBe(true));
      for (const outcome of await pending) {
        expect(outcome.status).toBe('rejected');
        if (outcome.status === 'rejected') expect(outcome.reason.message).toContain('document changed');
      }
      expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(false);
      await act(async () => { workers[1]!.release(); });
      expect(result.current.mainOpens).toEqual([]);
      expect(result.current.renderer.error).toBeNull();
      if (action === 'reload') unmount();
    });
  }
}

test('read-only handoff fallback requests a new frame and restores queries', async () => {
  installWorker({ failState: true });
  const { result } = renderHook(useHarness, { initialProps: { ...initialProps, readOnly: true } });
  await waitFor(() => expect(result.current.mainOpens).toEqual([true]));
  await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
  await waitFor(() => expect(result.current.renderer.frame).not.toBeNull());
  await waitFor(() => expect(result.current.renderer.queries?.isReady()).toBe(true));
  expect(result.current.renderer.queries!.pageCount()).toBeGreaterThan(0);
  expect(result.current.renderer.queries!.pageBounds(0)).not.toBeNull();
  expect(await result.current.renderer.resolveQueries()).toMatchObject({ queries: result.current.renderer.queries });
  expect(result.current.errors).toEqual([]);
});

test('read-only handoff OOM bootstraps its replacement worker and restores queries', async () => {
  const { workers, posted } = installWorker({ oomStage: 'encodeState' });
  const { result } = renderHook(useHarness, { initialProps: { ...initialProps, readOnly: true } });
  await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
  await waitFor(() => expect(workers).toHaveLength(2));
  await waitFor(() => expect(posted.filter((request) => request.type === 'bootstrap')).toHaveLength(2));
  await waitFor(() => expect(result.current.renderer.workerSurfacesActive).toBe(true));
  await waitFor(() => expect(result.current.renderer.queries?.isReady()).toBe(true));
  expect(result.current.renderer.frame).not.toBeNull();
  expect(result.current.renderer.queries!.pageCount()).toBeGreaterThan(0);
  expect(result.current.renderer.queries!.pageBounds(0)).not.toBeNull();
  expect(result.current.mainOpens).toEqual([false]);
  expect(result.current.errors).toEqual([]);
});

test('disabled worker open preserves main-thread open, projection and flush without replica helpers', async () => {
  const spies = [
    spyOn(replicaHelpers, 'deferWorkerOpenReplica'),
    spyOn(replicaHelpers, 'ensureWorkerOpenReplica'),
    spyOn(replicaHelpers, 'awaitWorkerOpenReplica'),
    spyOn(replicaHelpers, 'workerOpenReplicaPending'),
    spyOn(replicaHelpers, 'workerOpenSourceVersion'),
    spyOn(replicaHelpers, 'failWorkerOpenReplica'),
  ];
  try {
    const { workers, posted } = installWorker();
    const replicas: Array<YrsSession | null> = [];
    const props: HarnessProps = {
      ...initialProps,
      experimentalWorkerOpen: false,
      collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    };
    const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    expect(result.current.mainOpens).toEqual([true]);
    expect(replicas).toEqual([session]);
    expect(result.current.loadChecks).toEqual([1, 1, 1]);
    act(() => rerender({ ...props }));
    await act(async () => {});
    expect(replicas).toEqual([session]);
    expect(result.current.loadChecks).toEqual([1, 1, 1]);
    expect(result.current.core.documentFromYrs()).not.toBeNull();
    act(() => result.current.core.scheduleCompatibilityWarm());
    result.current.core.cancelCompatibilityWarm();
    expect(await result.current.ref.current!.readParagraphs({ view: 'accepted' })).toMatchObject({ ok: true });
    await result.current.ref.current!.flushPendingInput();
    expect(await result.current.bridgeRef.current!.runAfterPendingInput(() => true)).toBe(true);
    expect(result.current.bridgeRef.current!.hasPendingInput()).toBe(false);
    const paragraph = session.paragraphs('body')[0]!;
    act(() => session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited '));
    expect(result.current.core.documentFromYrs()).not.toBeNull();
    act(() => result.current.pipeline.runLayoutPipeline({ onHost: true }));
    await waitFor(() => expect(result.current.renderer.frame).not.toBeNull());
    unmount();
    expect(replicas).toEqual([session, null]);
    for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    expect(posted.some((request) => ['open', 'fontRequirements', 'encodeState'].includes(request.type))).toBe(false);
    expect(workers).toHaveLength(1);
  } finally {
    for (const spy of spies) spy.mockRestore();
  }
});
