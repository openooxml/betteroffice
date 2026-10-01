import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createRef, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import {
  createYrsSession,
  preloadResidentEngineWorker,
  ResidentWorkerOutOfMemoryError,
  proposalRevisionPreview,
  createYrsPositionProjection,
  yrsLocToProjectedDisplayPosition,
  type YrsDocxHost,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import {
  residentWorkerFactory,
  type InProcessResidentWorker,
} from '@betteroffice/docx/yrs/__fixtures__/residentWorker';
import type {
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { LayoutSelectionGate } from '@betteroffice/docx/layout';
import { decodeFrameDelta } from '@betteroffice/docx/layout/render';
import { useCanvasRenderer, type OpenInWorker } from './useDisplayList';
import { useLayoutPipeline } from './useLayoutPipeline';
import { useHostSearch, type DocxSearchState } from './useHostSearch';
import { useYrsCoreSession } from './useYrsCoreSession';
import type { DocxEditorCollaborationOptions } from '../types';
import { awaitWorkerOpenReplica, ensureWorkerOpenReplica, requestWorkerOpenReplica } from '../internals/workerOpenReplica';
import { isLayoutQueued, revisionPreviewKey, revisionPreviewKeyOf, sourceVersionOf } from '../internals/layoutProvenance';
import * as replicaHelpers from '../internals/workerOpenReplica';
import { registeredWorkerProposalAuthority, workerProposalAuthority } from '../internals/workerProposalAuthority';
import type { DocxEditorRef } from '../../DocxEditor';
import { PagedEditor, type PagedEditorRef } from '../PagedEditor';
import { UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import { createCommentIdAllocator } from '../commentFactories';
import { DocxReplicaNotReadyError, useDocxEditorRefApi } from './useDocxEditorRefApi';
import { usePagedEditorCommandBridge, type PagedEditorCommandBridge } from './usePagedEditorRefApi';
import { YrsInput, type YrsInputRef } from '../YrsInput';
import { flushEditorInput } from '../editorBatches';

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
const editModule = new WebAssembly.Module(
  new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
);
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>>;

beforeEach(() => {
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(editModule);
});

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});
afterEach(() => {
  cleanup();
  compileModule.mockRestore();
  globalThis.Worker = originalWorker;
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  await act(async () => {});
  if (ownsDom) await GlobalRegistrator.unregister();
});

function installWorker(options: {
  failOpen?: boolean;
  failState?: boolean;
  failProposal?: boolean;
  crashProposalOnce?: boolean;
  /** Fails the replacement worker's first snapshot once this settles. */
  failReplacementSnapshot?: Promise<void>;
  holdState?: boolean;
  holdOpen?: boolean;
  oomStage?: 'open' | 'fontRequirements' | 'bootstrap' | 'encodeState' | 'proposal' |
    'documentRead';
  holdRetryOpen?: boolean;
  revisionCount?: number;
  failRevisionCount?: boolean;
  onRevisionCount?: () => void;
  holdBootstrap?: boolean;
  holdSync?: boolean;
  holdCompletion?: boolean;
  holdReply?: (request: ResidentEngineWorkerRequest) => boolean;
  /** Decision replies arrive without the font requirements of the host's next pass. */
  withoutDecisionFontRequirements?: boolean;
  refusePreview?: boolean;
} = {}) {
  const workers: InProcessResidentWorker[] = [];
  const posted: ResidentEngineWorkerRequest[] = [];
  const hostModules: ResidentEngineWorkerHostModule[] = [];
  const replies = new Map<number, () => void>();
  const received = new Set<ResidentEngineWorkerRequest>();
  const responses = new Map<ResidentEngineWorkerRequest, ResidentEngineWorkerResponse>();
  const replyWaiters = new Set<() => void>();
  let proposalCrashed = false;
  let snapshotFailed = false;
  globalThis.Worker = class {
    constructor() {
      const worker = startWorker();
      const requests = new Map<number, ResidentEngineWorkerRequest>();
      let onmessage: typeof worker.onmessage = null;
      if (options.holdReply) Object.defineProperty(worker, 'onmessage', {
        get: () => {
          const listener = onmessage;
          return (event: MessageEvent<ResidentEngineWorkerResponse>) => {
            const request = posted.find((request) => request.id === event.data.id);
            const data = event.data;
            const delivered = options.withoutDecisionFontRequirements && data.ok && data.proposal &&
              request?.type === 'proposal' && request.operation.kind === 'setStates'
              ? { data: { ...data, proposal: { ...data.proposal, fontRequirements: undefined } } } as typeof event
              : event;
            const deliver = () => listener?.(delivered);
            if (request && options.holdReply?.(request)) replies.set(request.id, deliver);
            else deliver();
            if (request) received.add(request);
            const workerRequest = requests.get(event.data.id);
            if (workerRequest) responses.set(workerRequest, event.data);
            for (const waiter of replyWaiters) waiter();
          };
        },
        set: (listener: typeof worker.onmessage) => { onmessage = listener; },
      });
      const send = worker.postMessage.bind(worker);
      worker.postMessage = (request, transfer) => {
        if (request.type === 'editModule') {
          hostModules.push(request);
          send(request, transfer);
          return;
        }
        posted.push(request);
        requests.set(request.id, request);
        if ((options.holdState && request.type === 'encodeState') ||
            (options.holdOpen && request.type === 'open') ||
            (options.holdBootstrap && request.type === 'bootstrap') ||
            (options.holdSync && request.type === 'sync') ||
            (options.holdRetryOpen && workers.length > 1 && request.type === 'open') ||
            (options.holdCompletion && request.type === 'completeLayout')) worker.hold();
        if (options.oomStage === request.type &&
            (!['encodeState', 'proposal', 'documentRead'].includes(request.type) ||
              workers.length === 1)) {
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'worker exhausted memory', terminal: true, outOfMemory: true },
          } as MessageEvent));
        } else if (options.crashProposalOnce && !proposalCrashed &&
            request.type === 'proposal' && request.operation.kind === 'propose') {
          proposalCrashed = true;
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'proposal crashed', terminal: true },
          } as MessageEvent));
        } else if (options.failReplacementSnapshot && !snapshotFailed && workers.indexOf(worker) > 0 &&
            request.type === 'proposal' && request.operation.kind === 'snapshot') {
          snapshotFailed = true;
          void options.failReplacementSnapshot.then(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'snapshot failed' },
          } as MessageEvent));
        } else if (options.refusePreview && request.type === 'open' && request.previewBlocks !== undefined) {
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: true, previewRefused: true },
          } as MessageEvent));
        } else if ((options.failOpen && request.type === 'open') ||
            (options.failState && request.type === 'encodeState')) {
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'open failed', terminal: true },
          } as MessageEvent));
        } else if (options.failProposal && request.type === 'proposal' && request.operation.kind === 'propose') {
          queueMicrotask(() => worker.onmessage?.({
            data: { id: request.id, ok: false, error: 'proposal failed' },
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
  return {
    workers, posted, hostModules, replies, responses,
    received(type: ResidentEngineWorkerRequest['type'], afterId = 0): Promise<ResidentEngineWorkerRequest> {
      return new Promise((resolve) => {
        const check = () => {
          const request = [...received].find((request) => request.type === type && request.id > afterId);
          if (!request) return;
          replyWaiters.delete(check);
          resolve(request);
        };
        replyWaiters.add(check);
        check();
      });
    },
    reply(request: ResidentEngineWorkerRequest) {
      const deliver = replies.get(request.id);
      if (!deliver) throw new Error(`No held reply for ${request.type} ${request.id}`);
      replies.delete(request.id);
      deliver();
    },
  };
}

interface HarnessProps {
  experimentalWorkerOpen: boolean;
  hydrateOnDemand?: boolean;
  previewFirstPage?: boolean;
  /** Opens the first-page preview in the worker, as DocxEditor does. */
  workerPreview?: boolean;
  openInWorker?: OpenInWorker;
  source: Uint8Array;
  generation: number;
  collaboration?: DocxEditorCollaborationOptions;
  readOnly?: boolean;
  workerProposals?: boolean;
  onWorkerRevisions?: () => void;
  onWorkerContentChange?: () => void;
  allowHostProposals?: boolean;
  resolvedCommentIds?: ReadonlySet<number>;
  measurementFont?: Uint8Array;
  onHostDocument?: (session: YrsSession) => void;
  onLoad?: (api: DocxEditorRef) => void;
  onPresented?: (session: unknown) => void;
  /** Asks for the replica as soon as the session exists, as DocxEditor does for plugins, sidebars or the outline. */
  wanted?: boolean;
  /** Passes the renderer's own pending completion, as DocxEditor does. */
  followCompletion?: boolean;
  /** Holds the replica as while the shown engine's completion is still to be asked of the worker. */
  holdReplica?: boolean;
}

function useHarness(props: HarnessProps) {
  const relayout = useRef<(() => void) | null>(null);
  const workerRelayout = useRef<(() => void) | null>(null);
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
  const [commentsSidebarOpen, setCommentsSidebarOpen] = useState(false);
  const mainOpens = useRef<boolean[]>([]);
  const errors = useRef<Error[]>([]);
  const notifiedErrors = useRef(new WeakSet<Error>());
  const notifyError = useCallback((error: Error) => {
    if (notifiedErrors.current.has(error)) return;
    notifiedErrors.current.add(error);
    errors.current.push(error);
  }, []);
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
      onHostDocument: (host, _generation, session) => {
        setHost(host);
        props.onHostDocument?.(session);
      },
      onError: notifyError,
    },
    {
      previewFirstPage: props.previewFirstPage,
      heldEngine: renderer.layoutEngine,
      shownEngine: renderer.presentedEngine,
      workerOpen: props.experimentalWorkerOpen ? {
        openInWorker,
        ...(props.workerPreview ? { openPreviewInWorker: renderer.openPreviewInWorker } : {}),
        workerProposals: props.workerProposals,
        refreshWorkerLayout: () => workerRelayout.current?.(),
        renderedFrame: renderer.status === 'ready' ? renderer.displayList : null,
        ...(props.holdReplica ? { pendingCompletion: renderer.presentedEngine } : {}),
        ...(props.followCompletion ? { pendingCompletion: renderer.pendingCompletion } : {}),
        hydrateOnDemand: props.hydrateOnDemand,
        onWorkerRevisions: props.onWorkerRevisions,
        onWorkerContentChange: props.onWorkerContentChange,
      } : undefined,
    }
  );
  useEffect(() => {
    const error = renderer.error;
    const owner = renderer.errorEngine;
    if (!error || (props.experimentalWorkerOpen && owner && owner !== core.session &&
      (owner as YrsSession).isDisplayOnly?.() !== true)) return;
    if (!core.failOpening(error, owner ?? undefined)) notifyError(error);
  }, [core.failOpening, core.session, notifyError, props.experimentalWorkerOpen, renderer.error, renderer.errorEngine]);
  useEffect(() => {
    if (renderer.status === 'ready') props.onPresented?.(renderer.presentedEngine);
  }, [props.onPresented, renderer.presentedEngine, renderer.status]);
  handoffFromRef.current = core.handoffFrom;
  const replicaPending = Boolean(core.hydrateOnDemand && core.session && !core.replicaReady);
  useEffect(() => {
    if ((props.wanted || commentsSidebarOpen) && replicaPending) core.requestReplica();
  }, [commentsSidebarOpen, core, props.wanted, replicaPending]);
  const syncCoordinator = useRef(new LayoutSelectionGate());
  const element = useRef<HTMLDivElement | null>(null);
  const registeredFont = useRef<{ session: YrsSession; bytes: Uint8Array; id: number } | null>(null);
  const pipeline = useLayoutPipeline({
    document: host?.document ?? null,
    session: core.session,
    renderEnv: {},
    pageGap: 24,
    zoom: 1,
    residentMeasurementConfig: (requirements) => {
      const session = core.session;
      if (!session) return null;
      const fontBytes = props.measurementFont ?? font;
      if (registeredFont.current?.session !== session || registeredFont.current.bytes !== fontBytes) {
        registeredFont.current = { session, bytes: fontBytes, id: session.registerFont(fontBytes) };
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
      if (!core.failOpening(error, session)) notifyError(error);
    },
  });
  relayout.current = pipeline.runLayoutPipeline;
  workerRelayout.current = () => pipeline.scheduleLayout('remote', true);
  useEffect(() => {
    if (props.readOnly && core.session) relayout.current?.();
  }, [props.readOnly, core.session]);
  const pagedEditorRef = useRef<PagedEditorRef | null>(null);
  const coreRef = useRef(core);
  coreRef.current = core;
  const searchReveals = useRef<number[]>([]);
  pagedEditorRef.current = useMemo(() => core.session ? {
    getYrsSession: () => coreRef.current.session,
    getDocument: () => coreRef.current.documentFromYrs(),
    hasPendingInput: () => false,
    flushPendingInput: async () => {},
    yrsLocToDisplayPosition: (loc: Parameters<PagedEditorRef['yrsLocToDisplayPosition']>[0]) => {
      const session = coreRef.current.session!;
      const projection = createYrsPositionProjection(session, 'body');
      return yrsLocToProjectedDisplayPosition(session, () => projection, loc);
    },
    revealDisplayPosition: (position: number) => { searchReveals.current.push(position); return 'scrolled'; },
    syncYrsInputState: () => true,
    refreshWorkerLayout: () => workerRelayout.current?.(),
  } as unknown as PagedEditorRef : null, [core.session]);
  const hostSearch = useHostSearch({
    pagedEditorRef,
    displayListQueries: renderer.queries,
    canvasHostRef: element,
  });
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
    allowHostProposalsRef: { current: props.allowHostProposals === true },
    hostSearch: hostSearch.api,
    settledDisplayList: renderer.settledDisplayList,
  });
  const loaded = useRef(new WeakSet<YrsSession>());
  useEffect(() => {
    if (!host || !core.session || !ref.current || loaded.current.has(core.session)) return;
    loaded.current.add(core.session);
    props.onLoad?.(ref.current);
  }, [core.session, host, props.onLoad]);
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
    core, renderer, pipeline, host, ref, bridgeRef, pagedEditorRef,
    searchHighlight: hostSearch.highlight,
    searchReveals: searchReveals.current,
    openCommentsSidebar: () => setCommentsSidebarOpen(true),
    presentFrame: () => core.notifyFramePresented(renderer.presentedEngine),
    loadChecks: loadChecks.current, mainOpens: mainOpens.current, errors: errors.current,
  };
}

const initialProps: HarnessProps = { experimentalWorkerOpen: true, source: bytes, generation: 1 };

test('eager worker open preserves input and command order after first paint until loadState completes', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const source = await longFixture(1);
  const editor = createRef<PagedEditorRef>();
  const bridge = { current: null as PagedEditorCommandBridge | null };
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function Editable() {
    harness = useHarness({ ...initialProps, source, hydrateOnDemand: false });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries}
        commandBridgeRef={bridge} />
    </>;
  }
  const view = render(<Editable />);
  await waitFor(() => expect(harness.host).not.toBeNull());
  act(() => harness.pipeline.runLayoutPipeline());
  await waitFor(() => expect(harness.renderer.status).toBe('ready'));
  act(() => harness.presentFrame());
  await waitFor(() => expect(posted.some((r) => r.type === 'encodeState')).toBe(true));
  const session = harness.core.session!;
  const load = spyOn(session, 'loadState');
  const insert = spyOn(session, 'insertText');
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  expect(textarea.readOnly).toBe(false);
  await waitFor(() => expect(document.activeElement).toBe(textarea));
  const queries = harness.renderer.queries!;
  const caret = queries.caretRect(6)!;
  const size = queries.pageSize(0)!;
  const canvas = canvasHost.current!.firstElementChild!;
  canvas.getBoundingClientRect = () => ({ left: 0, top: 0, right: size.width,
    bottom: size.height, ...size }) as DOMRect;
  fireEvent.mouseDown(canvas, { clientX: caret.x, clientY: caret.y + caret.height / 2, button: 0 });
  fireEvent.mouseUp(window, { clientX: caret.x, clientY: caret.y + caret.height / 2, button: 0 });
  fireEvent.input(textarea, { target: { value: 'A' } });
  let seen: string[] = [];
  const command = bridge.current!.runAfterPendingInput(() => {
    seen = session.paragraphs('body').map((p) => p.text);
  });
  fireEvent.input(textarea, { target: { value: 'B' } });
  fireEvent.keyDown(textarea, { key: 'Enter' });
  fireEvent.paste(textarea, { clipboardData: { getData: () => 'P\r\nQ' } });
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await act(async () => { await Promise.resolve(); });
  fireEvent.input(textarea);
  fireEvent.keyDown(textarea, { key: 'Backspace' });
  const flush = editor.current!.flushPendingInput();
  await act(async () => { await Promise.resolve(); });
  expect(editor.current!.hasPendingInput()).toBe(true);
  expect(insert).not.toHaveBeenCalled();
  expect(load).not.toHaveBeenCalled();
  expect(seen).toEqual([]);
  await act(async () => { workers[0].release(); await command; await flush; });
  expect(load).toHaveBeenCalledTimes(1);
  expect(seen[0]).toBe('FirstA paragraph');
  expect(session.paragraphs('body').slice(0, 3).map((p) => p.text))
    .toEqual(['FirstAB', 'P', 'Q日 paragraph']);
  expect(session.selection()?.head.offset).toBe(2);
  expect(textarea.value).toBe('');
  load.mockRestore();
  insert.mockRestore();
});

test('read-only on-demand worker open supersedes pending select-all when admitting a command', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const source = await longFixture(2);
  const editor = createRef<PagedEditorRef>();
  const bridge = { current: null as PagedEditorCommandBridge | null };
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function ReadOnly() {
    harness = useHarness({ ...initialProps, source, readOnly: true, hydrateOnDemand: true });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core} readOnly
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries}
        commandBridgeRef={bridge} />
    </>;
  }
  const view = render(<ReadOnly />);
  await waitFor(() => expect(harness.renderer.status).toBe('ready'));
  act(() => harness.presentFrame());
  const session = harness.core.session!;
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  expect(textarea.readOnly).toBe(true);
  expect(harness.core.replicaReady).toBe(false);
  fireEvent.keyDown(textarea, { key: 'a', ctrlKey: true });
  await waitFor(() => expect(posted.some((r) => r.type === 'encodeState')).toBe(true));
  const command = bridge.current!.runAfterPendingInput(() => session.selection());
  let selected!: ReturnType<YrsSession['selection']>;
  await act(async () => { workers[0].release(); selected = await command; });
  const paragraphs = session.paragraphs('body');
  expect(selected).not.toEqual({
    anchor: { story: 'body', paraId: paragraphs[0].paraId, offset: 0 },
    head: { story: 'body', paraId: paragraphs.at(-1)!.paraId, offset: paragraphs.at(-1)!.text.length },
  });
});

test('eager worker-open hydration failure rejects flush, command and save during composition', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const source = await longFixture(1);
  const editor = createRef<PagedEditorRef>();
  const bridge = { current: null as PagedEditorCommandBridge | null };
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function Editable() {
    harness = useHarness({ ...initialProps, source, hydrateOnDemand: false });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries}
        commandBridgeRef={bridge} />
    </>;
  }
  const view = render(<Editable />);
  await waitFor(() => expect(harness.host).not.toBeNull());
  act(() => harness.pipeline.runLayoutPipeline());
  await waitFor(() => expect(harness.renderer.status).toBe('ready'));
  act(() => harness.presentFrame());
  await waitFor(() => expect(posted.some((r) => r.type === 'encodeState')).toBe(true));
  const session = harness.core.session!;
  const failure = new Error('hydration failed');
  const open = spyOn(session, 'openDocx').mockImplementation(() => { throw failure; });
  const insert = spyOn(session, 'insertText');
  const operation = mock(() => true);
  try {
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    expect(textarea.readOnly).toBe(false);
    expect(harness.core.replicaReady).toBe(false);
    fireEvent.compositionStart(textarea);
    textarea.value = '日本';
    const calls = Promise.allSettled([
      editor.current!.flushPendingInput(),
      bridge.current!.runAfterPendingInput(operation),
      flushEditorInput(editor, true),
    ]);
    const request = posted.find((r) => r.type === 'encodeState')!;
    await act(async () => {
      workers[0].onmessage?.({
        data: { id: request.id, ok: false, error: failure.message },
      } as MessageEvent<ResidentEngineWorkerResponse>);
    });
    const settled = await calls;
    expect(settled[0]).toEqual({ status: 'rejected', reason: failure });
    expect(settled[1]).toEqual({ status: 'rejected', reason: failure });
    expect(settled[2]).toEqual({
      status: 'fulfilled', value: { ok: false, code: 'input-failed', error: failure },
    });
    expect(editor.current!.hasPendingInput()).toBe(true);
    expect(textarea.value).toBe('日本');
    expect(operation).not.toHaveBeenCalled();
    expect(insert).not.toHaveBeenCalled();
    expect(session.storyIds()).toEqual([]);
  } finally {
    open.mockRestore();
    insert.mockRestore();
  }
}, 3_000);

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
  const run = () => {
    const pending = [...frames.values()];
    frames.clear();
    for (const callback of pending) callback(performance.now());
  };
  return {
    run,
    async until<T>(promise: Promise<T>): Promise<T> {
      let settled = false;
      void promise.then(() => { settled = true; }, () => { settled = true; });
      while (!settled) {
        run();
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      return promise;
    },
    async untilCommitted<T>(promise: Promise<T>): Promise<T> {
      let settled = false;
      void promise.then(() => { settled = true; }, () => { settled = true; });
      while (!settled) {
        await act(async () => {
          run();
          await new Promise<void>((resolve) => setImmediate(resolve));
        });
      }
      await act(async () => {});
      return promise;
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
    expect(rejection).toBeInstanceOf(Error);
    expect((rejection as Error).message).toContain('document changed');
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

test('a terminal worker failure before the accepted full frame tears down opening and notifies once', async () => {
  const options: Parameters<typeof installWorker>[0] = { holdReply: () => false };
  const { workers, posted, received } = installWorker(options);
  const frames = holdFrames();
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  let acceptFull!: (session: YrsSession) => void;
  const fullAccepted = new Promise<YrsSession>((resolve) => { acceptFull = resolve; });
  let presentPreview!: () => void;
  const previewPresented = new Promise<void>((resolve) => { presentPreview = resolve; });
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, previewFirstPage: true, source: longBytes,
      workerProposals: true, hydrateOnDemand: true,
      onHostDocument: (session) => { if (!session.isDisplayOnly()) acceptFull(session); },
      onPresented: (session) => { if ((session as YrsSession)?.isDisplayOnly()) presentPreview(); },
    },
  });
  try {
    await act(async () => { await received('open'); });
    await act(async () => {});
    expect(result.current.core.previewing).toBe(true);
    const preview = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await act(async () => {});
    await previewPresented;
    await act(async () => {});
    expect(result.current.renderer.presentedEngine).toBe(preview);
    act(() => result.current.presentFrame());
    await act(async () => frames.run());
    await act(async () => frames.run());
    await act(async () => { await fullAccepted; });
    expect(result.current.core.previewing).toBe(false);
    const full = result.current.core.session!;
    expect(result.current.core.handoffFrom).toBe(preview);
    const pending = awaitWorkerOpenReplica(full)!;
    const rejected = pending.catch((error: unknown) => error);
    options.oomStage = 'fontRequirements';
    await act(async () => {
      result.current.pipeline.runLayoutPipeline();
      await received('fontRequirements');
    });
    await act(async () => { await rejected; });
    expect(result.current.core.session).toBeNull();
    expect(result.current.core.handoffFrom).toBeNull();
    expect(result.current.core.opening).toBe(false);
    expect(result.current.errors).toEqual([result.current.renderer.error!]);
    expect(result.current.errors[0]).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(result.current.errors).toHaveLength(1);
    expect(workers).toHaveLength(2);
    for (let frame = 0; frame < 5; frame += 1) await act(async () => frames.run());
    expect(result.current.errors).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  } finally {
    unmount();
    errorLog.mockRestore();
    frames.restore();
  }
});

test('a preloaded spare worker takes the open that starts alongside the preview', async () => {
  const { workers, posted, hostModules } = installWorker();
  await preloadResidentEngineWorker();
  expect(workers).toHaveLength(1);
  expect(posted).toEqual([{ id: 1, type: 'warm', hostModule: true }]);
  expect(hostModules).toEqual([{ type: 'editModule', module: editModule }]);
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

test('a preview the worker opens lays out there, and the full open queues right behind its layout', async () => {
  const fullOpen = (request: ResidentEngineWorkerRequest) =>
    request.type === 'open' && request.previewBlocks === undefined;
  const { workers, posted, reply } = installWorker({ holdReply: fullOpen });
  const frames = holdFrames();
  try {
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, previewFirstPage: true, workerPreview: true, source: longBytes },
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    expect(preview.isDisplayOnly()).toBe(true);
    // This thread parsed none of it: the worker holds the preview.
    expect(preview.storyIds()).toEqual([]);
    expect(posted.map((request) => request.type)).toEqual(['open']);
    expect(posted[0].type === 'open' && posted[0].previewBlocks).toBeGreaterThan(0);

    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    expect(result.current.renderer.status).toBe('ready');
    expect(result.current.renderer.displayList?.pages.length).toBeGreaterThan(0);
    // The preview paints here: the full open runs in that worker next.
    expect(result.current.renderer.workerSurfacesActive).toBe(false);
    await waitFor(() => expect(posted.filter(fullOpen)).toHaveLength(1));
    const bootstrap = posted.findIndex((request) => request.type === 'bootstrap');
    expect(bootstrap).toBeGreaterThan(0);
    expect(posted[bootstrap].type === 'bootstrap' && posted[bootstrap].opened).toBe(true);
    expect(posted[bootstrap].type === 'bootstrap' && posted[bootstrap].displayWindow).toEqual([0, 2]);
    expect(posted.findIndex(fullOpen)).toBeGreaterThan(bootstrap);
    expect(posted.map((request) => request.type)).not.toContain('encodeState');
    expect(preview.storyIds()).toEqual([]);
    expect(result.current.mainOpens).toEqual([]);
    // A relayout of the preview after the full open took its worker over keeps the shown
    // layout: it asks nothing of the worker and opens nothing here.
    const shownLayout = result.current.pipeline.layout;
    const before = posted.length;
    act(() => result.current.pipeline.runLayoutPipeline());
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
    expect(posted.length).toBe(before);
    expect(result.current.pipeline.layout).toBe(shownLayout);
    expect(preview.storyIds()).toEqual([]);
    expect(result.current.errors).toEqual([]);

    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await act(async () => {
      reply(posted.find(fullOpen)!);
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(false));
    const full = result.current.core.session!;
    expect(full).not.toBe(preview);
    expect(full.isDisplayOnly()).toBe(false);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
    expect(workers).toHaveLength(1);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(2);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test('a preview font preflight answered after the full open took its worker over reports no error', async () => {
  const fullOpen = (request: ResidentEngineWorkerRequest) =>
    request.type === 'open' && request.previewBlocks === undefined;
  let holdRequirements = false;
  const held: ResidentEngineWorkerRequest[] = [];
  const { posted, reply } = installWorker({
    holdReply: (request) => {
      if (!holdRequirements || request.type !== 'fontRequirements') return false;
      held.push(request);
      return true;
    },
  });
  let openFull!: () => void;
  const fullOpenGate = new Promise<void>((resolve) => { openFull = resolve; });
  const frames = holdFrames();
  try {
    const openInWorker: OpenInWorker = async (...args) => {
      await fullOpenGate;
      return result.current.renderer.openInWorker(...args);
    };
    const { result, unmount } = renderHook(useHarness, {
      initialProps: {
        ...initialProps,
        previewFirstPage: true,
        workerPreview: true,
        source: longBytes,
        openInWorker,
      },
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    const shownLayout = result.current.pipeline.layout;
    expect(posted.filter(fullOpen)).toHaveLength(0);

    holdRequirements = true;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(held).toHaveLength(1));
    holdRequirements = false;
    openFull();
    await waitFor(() => expect(posted.filter(fullOpen)).toHaveLength(1));
    expect(posted.indexOf(held[0]!)).toBeLessThan(posted.findIndex(fullOpen));
    await act(async () => {
      reply(held[0]!);
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
    expect(result.current.errors).toEqual([]);
    expect(result.current.pipeline.layout).toBe(shownLayout);

    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(result.current.core.previewing).toBe(false));
    const full = result.current.core.session!;
    expect(full).not.toBe(preview);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
    expect(result.current.renderer.displayList?.pages.length).toBeGreaterThan(0);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test('a worker preview loaded here before its first layout stops using the worker', async () => {
  const { workers, posted } = installWorker();
  const frames = holdFrames();
  try {
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, previewFirstPage: true, workerPreview: true, source: longBytes },
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    expect(preview.storyIds()).toEqual([]);
    // Something on this thread needs the preview's content before its first layout.
    await act(async () => {
      await requestWorkerOpenReplica(preview);
    });
    expect(preview.storyIds()).not.toEqual([]);
    expect(preview.isDisplayOnly()).toBe(true);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    expect(posted.map((request) => request.type)).not.toContain('bootstrap');
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(result.current.core.previewing).toBe(false));
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() =>
      expect(result.current.renderer.presentedEngine).toBe(result.current.core.session)
    );
    // The preview's worker went with it; the full document opened in a new one.
    expect(workers).toHaveLength(2);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    frames.restore();
  }
});

test('a package the worker cannot preview opens its preview here and the full document in that worker', async () => {
  const { workers, posted } = installWorker({ refusePreview: true });
  const frames = holdFrames();
  try {
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, previewFirstPage: true, workerPreview: true, source: longBytes },
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    expect(preview.isDisplayOnly()).toBe(true);
    expect(preview.storyIds()).not.toEqual([]);
    await waitFor(() => expect(posted.filter((request) => request.type === 'open')).toHaveLength(2));
    expect(
      posted
        .filter((request) => request.type === 'open')
        .map((request) => request.type === 'open' && request.previewBlocks !== undefined)
    ).toEqual([true, false]);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(result.current.core.previewing).toBe(false));
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() =>
      expect(result.current.renderer.presentedEngine).toBe(result.current.core.session)
    );
    expect(workers).toHaveLength(1);
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
  const { workers, posted } = installWorker({ failState: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  act(() => result.current.pipeline.runLayoutPipeline());
  await waitFor(() => expect(result.current.mainOpens).toEqual([true]));
  await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
  expect(result.current.core.session?.hasStory('body')).toBe(true);
  expect(result.current.errors).toEqual([]);
  const bootstraps = posted.filter((request) => request.type === 'bootstrap').length;
  act(() => result.current.pipeline.scheduleLayout('remote', true));
  await waitFor(() => expect(result.current.renderer.frame).not.toBeNull());
  expect(workers).toHaveLength(1);
  expect(posted.filter((request) => request.type === 'bootstrap')).toHaveLength(bootstraps);
});

test('getEditorRef during load opens the main replica and recovers worker rendering', async () => {
  const { workers, posted, responses } = installWorker({ holdReply: () => false });
  const frames = holdFrames();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  let mainState: Uint8Array | null = null;
  let mirror: ReturnType<typeof spyOn<YrsSession, 'mirrorWorkerDocument'>> | undefined;
  let layoutHere: ReturnType<typeof spyOn<YrsSession, 'layoutDocumentWithRegionsRetainedJson'>> | undefined;
  let buildHere: ReturnType<typeof spyOn<YrsSession, 'buildDisplayListFrame'>> | undefined;
  let terminate: ReturnType<typeof spyOn<InProcessResidentWorker, 'terminate'>> | undefined;
  let session!: YrsSession;
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, workerProposals: true, allowHostProposals: true,
      onHostDocument: (next) => { session = next; },
      onLoad: (api) => {
        expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
        mirror = spyOn(session, 'mirrorWorkerDocument');
        layoutHere = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
        buildHere = spyOn(session, 'buildDisplayListFrame');
        terminate = spyOn(workers[0]!, 'terminate');
        expect(api.getEditorRef()?.getYrsSession()).toBe(session);
        expect(api.getEditorRef()?.getYrsSession()).toBe(session);
        mainState = session.encodeState();
      },
    },
  });
  try {
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    await waitFor(() => expect(mainState).not.toBeNull());
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([true]);
    expect(mirror).toHaveBeenCalledWith(null);
    expect(session.workerDocumentMirrored()).toBe(false);
    expect(terminate).toHaveBeenCalledTimes(1);
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, 3000));
    expect(workers).toHaveLength(2);
    expect(workers[1]!.requests[0]).toBe('bootstrap');
    const bootstrap = posted.find((request) => request.type === 'bootstrap')!;
    expect(bootstrap.type).toBe('bootstrap');
    if (bootstrap.type !== 'bootstrap') throw new Error('Missing bootstrap');
    expect(bootstrap.opened).toBeUndefined();
    expect(bootstrap.snapshot.workerAuthoritative).toBeUndefined();
    expect(bootstrap.snapshot.state).toEqual(mainState!);
    const reply = responses.get(bootstrap)!;
    expect(reply.ok).toBe(true);
    if (!reply.ok || !reply.frame) throw new Error('Missing bootstrap frame');
    expect(decodeFrameDelta(new Uint8Array(reply.frame)).full).toBe(true);
    expect(result.current.renderer.workerSurfacesActive).toBe(true);

    const api = result.current.ref.current!;
    const paragraph = (await api.getParagraphIdentities()).paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
    const proposalSyncs = posted.filter((request) => request.type === 'sync').length;
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: session.version(),
        proposals: [{
          id: 'recovered-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Recovered ',
        }],
      })).toMatchObject({ ok: true });
      result.current.pipeline.scheduleLayout('remote', true);
    });
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, 3000));
    expect(posted.filter((request) => request.type === 'sync').length).toBeGreaterThan(proposalSyncs);
    const proposed = posted.filter((request) => request.type === 'sync').at(-1)!;
    if (proposed.type !== 'sync') throw new Error('Missing proposal sync');
    expect(proposed.snapshot.workerAuthoritative).toBeUndefined();
    const proposal = await api.getProposals();
    const syncs = posted.filter((request) => request.type === 'sync').length;
    await act(async () => {
      expect(await api.setProposalStates({
        expectVersion: proposal.version, expectPreviewVersion: proposal.previewVersion,
        changes: [{ id: 'recovered-proposal', state: 'accepted' }],
      })).toMatchObject({ ok: true });
      result.current.pipeline.scheduleLayout('remote', true);
    });
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, 3000));
    expect(session.getProposals().proposals[0]!.state).toBe('accepted');
    expect(posted.filter((request) => request.type === 'sync').length).toBeGreaterThan(syncs);
    const accepted = posted.filter((request) => request.type === 'sync').at(-1)!;
    if (accepted.type !== 'sync') throw new Error('Missing proposal sync');
    expect(accepted.snapshot.workerAuthoritative).toBeUndefined();
    await waitFor(() => expect(sourceVersionOf(result.current.renderer.queries)).toBe(session.version()));
    expect(workers[1]!.requests).toContain('sync');
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([true]);
    expect(result.current.renderer.workerSurfacesActive).toBe(true);
    expect(layoutHere).not.toHaveBeenCalled();
    expect(buildHere).not.toHaveBeenCalled();
    expect(warning).toHaveBeenCalledTimes(1);
    expect(warning.mock.calls[0]![0]).toContain('[DocxEditor] getEditorRef()');
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    mirror?.mockRestore();
    layoutHere?.mockRestore();
    buildHere?.mockRestore();
    terminate?.mockRestore();
    warning.mockRestore();
    frames.restore();
  }
});

test('sync recovery follows the shown epoch and ignores a destroyed worker\'s late layout', async () => {
  let holdSync = false;
  const { workers, posted, responses, reply } = installWorker({
    holdState: true, holdReply: (request) => holdSync && request.type === 'sync',
  });
  const frames = holdFrames();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, hydrateOnDemand: true },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.frame).not.toBeNull());
    const shownEpoch = result.current.renderer.frame!.frameEpoch;
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
    const layoutHere = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
    const buildHere = spyOn(session, 'buildDisplayListFrame');
    const terminate = spyOn(workers[0]!, 'terminate');
    try {
      holdSync = true;
      act(() => result.current.pipeline.scheduleLayout('remote', true));
      act(() => frames.run());
      await waitFor(() => expect(posted.some((request) => request.type === 'sync')).toBe(true));
      const stale = posted.find((request) => request.type === 'sync')!;
      await waitFor(() => expect(responses.has(stale)).toBe(true));
      act(() => { void requestWorkerOpenReplica(session); });
      await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
      act(() => { result.current.ref.current!.getEditorRef(); });
      expect(terminate).toHaveBeenCalledTimes(1);
      holdSync = false;
      await frames.untilCommitted(result.current.renderer.settledDisplayList(null, 3000));
      expect(workers).toHaveLength(2);
      const bootstrap = posted.filter((request) => request.type === 'bootstrap').at(-1)!;
      if (bootstrap.type !== 'bootstrap') throw new Error('Missing recovery bootstrap');
      expect(bootstrap.expectedFrameEpoch).toBeGreaterThanOrEqual(shownEpoch);
      expect(bootstrap.opened).toBeUndefined();
      expect(bootstrap.snapshot.workerAuthoritative).toBeUndefined();
      const recovered = responses.get(bootstrap)!;
      if (!recovered.ok || !recovered.frame) throw new Error('Missing recovery frame');
      const delta = decodeFrameDelta(new Uint8Array(recovered.frame));
      expect(delta.full).toBe(true);
      expect(delta.frameEpoch).toBeGreaterThan(shownEpoch);
      const frame = result.current.renderer.frame;
      const layout = result.current.pipeline.layout;
      await act(async () => { reply(stale); workers[0]!.release(); });
      act(() => frames.run());
      expect(result.current.renderer.frame).toBe(frame);
      expect(result.current.pipeline.layout).toBe(layout);
      expect(result.current.mainOpens).toEqual([true]);
      expect(layoutHere).not.toHaveBeenCalled();
      expect(buildHere).not.toHaveBeenCalled();
      expect(warning).toHaveBeenCalledTimes(1);
      expect(result.current.errors).toEqual([]);
    } finally {
      layoutHere.mockRestore();
      buildHere.mockRestore();
      terminate.mockRestore();
    }
  } finally {
    unmount();
    warning.mockRestore();
    frames.restore();
  }
});

test('a failed synchronous main open keeps rendering pinned to the main engine', async () => {
  const { workers, posted } = installWorker();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, hydrateOnDemand: true },
  });
  let open: ReturnType<typeof spyOn<YrsSession, 'openDocx'>> | undefined;
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const failure = new Error('Main open failed');
    open = spyOn(session, 'openDocx').mockImplementation(() => { throw failure; });
    const ready = awaitWorkerOpenReplica(session)!;
    act(() => expect(() => result.current.ref.current!.getEditorRef()).toThrow(failure));
    await expect(ready).rejects.toBe(failure);
    expect(() => result.current.renderer.layoutInWorker(session, '{}')).toThrow(failure);
    expect(workers).toHaveLength(1);
    expect(posted.some((request) => request.type === 'bootstrap')).toBe(false);
    expect(result.current.errors).toEqual([failure]);
  } finally {
    open?.mockRestore();
    unmount();
    warning.mockRestore();
  }
});

test('opening the comments sidebar hydrates the pending replica and keeps worker layout', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames();
  const props = { ...initialProps, hydrateOnDemand: true, wanted: false };
  const { result, unmount } = renderHook(useHarness, { initialProps: props });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.frame).not.toBeNull());
    expect(result.current.core.replicaReady).toBe(false);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => result.current.openCommentsSidebar());
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => { workers[0]!.release(); await awaitWorkerOpenReplica(session); });
    expect(result.current.mainOpens).toEqual([false]);
    const layoutHere = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
    const buildHere = spyOn(session, 'buildDisplayListFrame');
    try {
      act(() => result.current.pipeline.scheduleLayout('remote', true));
      act(() => frames.run());
      await waitFor(() => expect(workers[0]!.requests).toContain('sync'));
      await act(async () => { await frames.until(result.current.renderer.settledDisplayList(null, null)); });
      expect(workers).toHaveLength(1);
      expect(workers[0]!.requests).toContain('sync');
      expect(result.current.renderer.workerSurfacesActive).toBe(true);
      expect(layoutHere).not.toHaveBeenCalled();
      expect(buildHere).not.toHaveBeenCalled();
      expect(result.current.errors).toEqual([]);
    } finally {
      layoutHere.mockRestore();
      buildHere.mockRestore();
    }
  } finally {
    unmount();
    frames.restore();
  }
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

const workerProposalProps: HarnessProps = {
  ...initialProps,
  source: longBytes,
  readOnly: true,
  hydrateOnDemand: true,
  workerProposals: true,
  allowHostProposals: true,
};

async function openWorkerProposals(props: HarnessProps = workerProposalProps) {
  const harness = renderHook(useHarness, { initialProps: props });
  await waitFor(() => expect(harness.result.current.core.workerProposalsReady).toBe(true));
  await act(async () => {
    await harness.result.current.ref.current!.whenLayoutComplete({ timeoutMs: 5000 });
  });
  return harness;
}

test('host search reads and navigates the resident worker without starting the main replica', async () => {
  const { workers, posted } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  const session = result.current.core.session!;
  const editor = result.current.pagedEditorRef.current!;
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const spies = [
    spyOn(replicaHelpers, 'awaitWorkerOpenReplica'),
    spyOn(replicaHelpers, 'ensureWorkerOpenReplica'),
    spyOn(replicaHelpers, 'requestWorkerOpenReplica'),
    spyOn(replicaHelpers, 'requestOnDemandWorkerOpenReplica'),
    spyOn(session, 'openDocx'),
    spyOn(session, 'loadState'),
    spyOn(session, 'searchText'),
    spyOn(session, 'encodeStickyPosition'),
    spyOn(session, 'resolveStickyPosition'),
    spyOn(editor, 'yrsLocToDisplayPosition'),
    spyOn(editor, 'flushPendingInput'),
    spyOn(workers[0], 'terminate'),
  ];
  try {
    const api = result.current.ref.current!;
    const events: Array<DocxSearchState | null> = [];
    const unsubscribe = api.onSearchChange((state) => events.push(state));
    await act(async () => {
      expect(await api.search('paragraph')).toEqual({
        query: 'paragraph', options: { caseSensitive: false }, total: 205, current: 0,
      });
    });
    let position = 1;
    const ranges = Array.from({ length: 205 }, (_, index) => {
      const text = index === 0 ? 'First paragraph' : index === 204 ? 'Tail paragraph' : `Paragraph ${index}`;
      const displayFrom = position + text.toLowerCase().indexOf('paragraph');
      position += text.length + 2;
      return { displayFrom, displayTo: displayFrom + 9 };
    });
    expect(result.current.searchHighlight?.matches.map(({ displayFrom, displayTo }) => ({
      displayFrom, displayTo,
    }))).toEqual(ranges);
    act(() => {
      expect(api.searchPrevious()?.current).toBe(204);
      expect(api.searchNext()?.current).toBe(0);
      expect(api.searchNext()?.current).toBe(1);
      expect(api.searchPrevious()?.current).toBe(0);
      expect(api.searchGoTo(206)?.current).toBe(1);
    });
    expect(api.getSearchState()).toMatchObject({ total: 205, current: 1 });
    expect(result.current.searchReveals).toEqual([
      ranges[0].displayFrom, ranges[204].displayFrom, ranges[0].displayFrom,
      ranges[1].displayFrom, ranges[0].displayFrom, ranges[1].displayFrom,
    ]);
    await act(async () => {
      expect(await api.search('paragraph', { caseSensitive: true })).toMatchObject({ total: 2, current: 0 });
    });
    act(() => api.clearSearch());
    expect(api.getSearchState()).toBeNull();
    expect(result.current.searchHighlight).toBeNull();
    expect(events.at(-1)).toBeNull();
    unsubscribe();
    expect(workerProposalAuthority(session)).not.toBeNull();
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.some(({ type }) => type === 'encodeState')).toBe(false);
    expect(posted.filter((request) => request.type === 'documentRead' && request.read.kind === 'searchText')).toHaveLength(2);
    for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    expect(warning.mock.calls.some(([message]) =>
      String(message).includes('needed the main-thread document')
    )).toBe(false);
  } finally {
    for (const spy of spies) spy.mockRestore();
    warning.mockRestore();
    unmount();
  }
}, 15_000);

test('host search refreshes through the worker and carries its current match after a proposal', async () => {
  const { workers, posted } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  const termination = spyOn(workers[0], 'terminate');
  const session = result.current.core.session!;
  const spies = [
    spyOn(session, 'searchText'), spyOn(session, 'encodeStickyPosition'),
    spyOn(session, 'resolveStickyPosition'), spyOn(result.current.pagedEditorRef.current!, 'yrsLocToDisplayPosition'),
  ];
  try {
    const api = result.current.ref.current!;
    const events: Array<DocxSearchState | null> = [];
    api.onSearchChange((state) => events.push(state));
    await act(async () => { await api.search('paragraph'); });
    act(() => api.searchGoTo(1));
    const before = result.current.searchHighlight!.matches[1].displayFrom;
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find(({ session }) => session?.story === 'body')!.session!;
    const initial = await api.getProposals();
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'search-edit', paragraph,
          suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'paragraph ',
        }],
      })).toMatchObject({ ok: true });
    });
    await waitFor(() => expect(result.current.searchHighlight?.matches).toHaveLength(206));
    expect(api.getSearchState()).toMatchObject({ total: 206, current: 2 });
    expect(result.current.searchHighlight!.matches[2].displayFrom).toBe(before + 10);
    expect(events.at(-1)).toMatchObject({ total: 206, current: 2 });
    const refreshes = posted.filter((request) => request.type === 'documentRead' && request.read.kind === 'searchText');
    expect(refreshes.length).toBeGreaterThan(1);
    expect(refreshes.at(-1)).toMatchObject({ read: { carry: { story: 'body' } } });
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.some(({ type }) => type === 'encodeState')).toBe(false);
    expect(termination).not.toHaveBeenCalled();
    for (const spy of spies) expect(spy).not.toHaveBeenCalled();
  } finally {
    for (const spy of spies) spy.mockRestore();
    termination.mockRestore();
    unmount();
  }
}, 15_000);

test('ASCII proposals skip font preflight and Unicode proposals request it before syncing', async () => {
  const { posted } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  try {
    const api = result.current.ref.current!;
    const identities = await api.getParagraphIdentities();
    const paragraphs = identities.paragraphs.filter(({ session }) => session?.story === 'body');
    const initialFonts = posted.filter(({ type }) => type === 'fontRequirements').length;
    const initialSyncs = posted.filter(({ type }) => type === 'sync').length;
    const propose = async (id: string, index: number, search: string, replaceWith: string) => {
      const current = await api.getProposals();
      await act(async () => {
        expect(await api.proposeChanges({ expectVersion: current.version, proposals: [{
          id, paragraph: paragraphs[index]!.session!,
          suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
          op: 'replaceText', search, replaceWith,
        }] })).toMatchObject({ ok: true });
      });
      await act(async () => { await api.whenLayoutComplete({ timeoutMs: 5000 }); });
    };
    await propose('ascii', 0, 'First', 'Leading');
    expect(posted.filter(({ type }) => type === 'sync').length).toBeGreaterThan(initialSyncs);
    expect(posted.filter(({ type }) => type === 'fontRequirements')).toHaveLength(initialFonts);
    const beforeUnicode = posted.length;
    await propose('unicode', 1, 'Paragraph', '漢字');
    expect(posted.filter(({ type }) => type === 'fontRequirements')).toHaveLength(initialFonts + 1);
    const followup = posted.slice(beforeUnicode).map(({ type }) => type);
    expect(followup.indexOf('fontRequirements')).toBeGreaterThan(followup.indexOf('proposal'));
    expect(followup.indexOf('sync')).toBeGreaterThan(followup.indexOf('fontRequirements'));
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('proposal font requirements are reused only for the same layout request', async () => {
  const { posted } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  try {
    const original = posted.find((request) => request.type === 'fontRequirements');
    if (original?.type !== 'fontRequirements') throw new Error('expected font preflight');
    const before = posted.filter(({ type }) => type === 'fontRequirements').length;
    const session = result.current.core.session!;
    const requirements = result.current.renderer.fontRequirementsInWorker;
    await act(async () => { await requirements(session, original.layoutInput); });
    expect(posted.filter(({ type }) => type === 'fontRequirements')).toHaveLength(before);
    const changed = JSON.stringify({
      ...JSON.parse(original.layoutInput),
      measurement: { defaults: { fontFamily: 'Courier New', fontSize: 11 } },
    });
    let answer!: string | null;
    await act(async () => { answer = await requirements(session, changed); });
    expect(posted.filter(({ type }) => type === 'fontRequirements')).toHaveLength(before + 1);
    expect(answer).toContain('Courier New');
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('decisions and their undos answer the font preflight from the decision reply', async () => {
  const { posted } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  try {
    const api = result.current.ref.current!;
    const session = result.current.core.session!;
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
    const initial = await api.getProposals();
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'decided', paragraph,
          suggest: { author: 'Host', date: '2026-10-01T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Added text ',
        }],
      })).toMatchObject({ ok: true });
    });
    await act(async () => { await api.whenLayoutComplete({ timeoutMs: 5000 }); });
    const fonts = () => posted.filter(({ type }) => type === 'fontRequirements').length;
    const before = fonts();
    for (const state of ['accepted', 'proposed', 'rejected', 'proposed'] as const) {
      const snapshot = await api.getProposals();
      const syncs = posted.filter(({ type }) => type === 'sync').length;
      await act(async () => {
        expect(await api.setProposalStates({
          expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
          changes: [{ id: 'decided', state }],
        })).toMatchObject({ ok: true });
      });
      await act(async () => { await api.whenLayoutComplete({ timeoutMs: 5000 }); });
      expect(posted.filter(({ type }) => type === 'sync').length).toBeGreaterThan(syncs);
      expect(revisionPreviewKeyOf(result.current.pipeline.layout))
        .toBe(revisionPreviewKey(proposalRevisionPreview(session.getProposals())));
    }
    expect(fonts()).toBe(before);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test.each(['A then B', 'B then A'])('Undo keeps worker proposals through font preflight replies %s', async (order) => {
  let holdRequirements = false;
  const { posted, replies, reply, received } = installWorker({
    holdReply: (request) => holdRequirements && request.type === 'fontRequirements',
    withoutDecisionFontRequirements: true,
  });
  const frames = holdFrames();
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  const { result, rerender, unmount } = renderHook(useHarness, {
    initialProps: { ...workerProposalProps, source: bytes },
  });
  let layoutHere: ReturnType<typeof spyOn> | undefined;
  let ensureReplica: ReturnType<typeof spyOn> | undefined;
  try {
    await act(async () => { await received('open'); });
    await act(async () => { await received('bootstrap'); });
    await act(async () => { await received('proposal'); });
    await act(async () => { await frames.until(result.current.renderer.settledDisplayList(null, null)); });
    const api = result.current.ref.current!;
    const session = result.current.core.session!;
    layoutHere = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
    ensureReplica = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
    const initial = await api.getProposals();
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'font-race-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Added text ',
        }],
      })).toMatchObject({ ok: true });
    });
    const settle = async () => {
      const previous = posted.filter((request) => request.type === 'sync').at(-1)?.id ?? 0;
      await act(async () => { frames.run(); await received('sync', previous); });
      await act(async () => { await frames.until(result.current.renderer.settledDisplayList(null, null)); });
      expect(isLayoutQueued(session)).toBe(false);
    };
    const setState = async (state: 'accepted' | 'proposed' | 'rejected') => {
      const snapshot = await api.getProposals();
      await act(async () => {
        expect(await api.setProposalStates({
          expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion,
          changes: [{ id: 'font-race-proposal', state }],
        })).toMatchObject({ ok: true });
      });
      expect(session.getProposals().proposals[0]!.state).toBe(state);
    };
    const assertCurrent = async () => {
      const preview = revisionPreviewKey(proposalRevisionPreview(session.getProposals()));
      expect(sourceVersionOf(result.current.pipeline.layout)).toBe(session.version());
      expect(revisionPreviewKeyOf(result.current.pipeline.layout)).toBe(preview);
      expect(sourceVersionOf(result.current.renderer.queries)).toBe(session.version());
      expect(revisionPreviewKeyOf(result.current.renderer.queries)).toBe(preview);
      const authority = workerProposalAuthority(session)!;
      expect(authority.holdsWorkerState()).toBe(true);
      const geometry = authority.geometry()!;
      expect(geometry.version).toBe(session.version());
      expect(geometry.previewVersion).toBe(session.getProposals().previewVersion);
      const geometryTarget = geometry.targets['font-race-proposal']!;
      expect(geometryTarget).toMatchObject({ ok: true });
      if (!geometryTarget.ok) throw new Error('proposal geometry unavailable');
      expect(geometryTarget.ranges.length).toBeGreaterThan(0);
      const queries = result.current.renderer.queries!;
      const rects = geometryTarget.ranges.flatMap(({ from, to }) => queries.rangeRects(from, to));
      expect(rects.length).toBeGreaterThan(0);
      for (const rect of rects) {
        expect(rect.width).toBeGreaterThan(0);
        expect(rect.height).toBeGreaterThan(0);
      }
      const target = await authority.navigationTarget(paragraph.story, paragraph.paraId, () => {
        throw new Error('unexpected main-thread navigation');
      });
      expect(target.version).toBe(session.version());
      expect(target.target).toMatchObject({ loc: { story: paragraph.story, paraId: paragraph.paraId } });
      if (!target.target || typeof target.target === 'string') throw new Error('proposal navigation unavailable');
      expect(target.target.position).toBeGreaterThan(0);
      expect(queries.caretRect(target.target.position)?.height).toBeGreaterThan(0);
      expect(layoutHere).not.toHaveBeenCalled();
      expect(ensureReplica).not.toHaveBeenCalled();
      expect(result.current.renderer.error).toBeNull();
      expect(result.current.errors).toEqual([]);
      expect(errorLog).not.toHaveBeenCalled();
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
      expect(result.current.mainOpens).toEqual([]);
    };
    await settle();
    await setState('accepted');
    await settle();
    holdRequirements = true;
    const previous = posted.filter((request) => request.type === 'fontRequirements').at(-1)!.id;
    await setState('proposed');
    let passA!: ResidentEngineWorkerRequest;
    await act(async () => { frames.run(); passA = await received('fontRequirements', previous); });
    expect(replies.has(passA.id)).toBe(true);
    const loadedFont = new Uint8Array(readFileSync(resolve(
      import.meta.dir, '../../../../../../crates/docx-raster/tests/assets/Carlito-Regular.ttf'
    )));
    rerender({ ...workerProposalProps, source: bytes, measurementFont: loadedFont });
    await act(async () => {
      result.current.pipeline.runLayoutPipeline();
      await new Promise((resolve) => setTimeout(resolve, 50));
    });
    // Pass B shares pass A's read while their inputs are identical, else it holds its own.
    const held = posted.filter((request) =>
      request.type === 'fontRequirements' && request.id >= passA.id && replies.has(request.id)
    );
    expect(held[0]).toBe(passA);
    expect(held.length).toBeLessThanOrEqual(2);
    holdRequirements = false;
    const previousSync = posted.filter((request) => request.type === 'sync').at(-1)!.id;
    await act(async () => {
      for (const pass of order === 'A then B' ? held : [...held].reverse()) reply(pass);
    });
    expect(layoutHere).not.toHaveBeenCalled();
    expect(ensureReplica).not.toHaveBeenCalled();
    expect(result.current.errors).toEqual([]);
    expect(errorLog).not.toHaveBeenCalled();
    await act(async () => { await received('sync', previousSync); });
    await act(async () => { await frames.until(result.current.renderer.settledDisplayList(null, null)); });
    expect(isLayoutQueued(session)).toBe(false);
    await assertCurrent();
    await setState('rejected');
    await settle();
    await setState('proposed');
    await settle();
    await assertCurrent();
  } finally {
    layoutHere?.mockRestore();
    ensureReplica?.mockRestore();
    unmount();
    errorLog.mockRestore();
    frames.restore();
  }
});

test.each(['unavailable', 'no adoption', 'no snapshot'])('a holding session with %s worker layout fails once without a frame loop', async (path) => {
  const { posted, received } = installWorker({ holdReply: () => false });
  const frames = holdFrames();
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...workerProposalProps, source: bytes },
  });
  try {
    await act(async () => { await received('open'); });
    await act(async () => { await received('bootstrap'); });
    await act(async () => { await received('proposal'); });
    await act(async () => { await frames.until(result.current.renderer.settledDisplayList(null, null)); });
    const session = result.current.core.session!;
    const api = result.current.ref.current!;
    const snapshot = await api.getProposals();
    await act(async () => {
      expect(await api.setProposalStates({
        expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion, changes: [],
      })).toMatchObject({ ok: true });
    });
    expect(workerProposalAuthority(session)?.holdsWorkerState()).toBe(true);
    const layoutHere = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
    const ensureReplica = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
    const requestFrame = spyOn(globalThis, 'requestAnimationFrame');
    const requirements = posted.filter((request) => request.type === 'fontRequirements');
    const invalid = path === 'no snapshot'
      ? spyOn(session, 'residentWorkerSnapshot').mockReturnValue(null)
      : path === 'unavailable'
        ? spyOn(session, 'isDisplayOnly').mockReturnValue(true)
        : null;
    const adopt = session.adoptResidentWorkerLayout;
    if (path === 'no adoption') session.adoptResidentWorkerLayout = undefined as never;
    try {
      act(() => result.current.pipeline.scheduleLayout('remote'));
      await act(async () => { frames.run(); });
      await waitFor(() => expect(result.current.renderer.error).not.toBeNull());
      const failure = result.current.renderer.error;
      expect(failure?.message).toBe('The resident worker holding proposals cannot lay out the document');
      expect(result.current.errors).toEqual([failure!]);
      const requests = posted.length;
      const frameRequests = requestFrame.mock.calls.length;
      for (let frame = 0; frame < 8; frame += 1) await act(async () => frames.run());
      expect(requestFrame.mock.calls.length).toBe(frameRequests);
      expect(posted).toHaveLength(requests);
      expect(posted.filter((request) => request.type === 'fontRequirements')).toHaveLength(requirements.length);
      expect(result.current.errors).toEqual([failure!]);
      expect(errorLog).toHaveBeenCalledTimes(1);
      expect(layoutHere).not.toHaveBeenCalled();
      expect(ensureReplica).not.toHaveBeenCalled();
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    } finally {
      session.adoptResidentWorkerLayout = adopt;
      invalid?.mockRestore();
      layoutHere.mockRestore();
      ensureReplica.mockRestore();
      requestFrame.mockRestore();
    }
  } finally {
    unmount();
    errorLog.mockRestore();
    frames.restore();
  }
});

test('a toggle and local refresh wait for an older worker sync without rebuilding proposals', async () => {
  const options = { holdReply: (request: ResidentEngineWorkerRequest) => request.type === 'sync' };
  const { workers, posted, replies, reply, received } = installWorker(options);
  const frames = holdFrames();
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...workerProposalProps, source: bytes },
  });
  try {
    await act(async () => { await received('open'); });
    await act(async () => { await received('bootstrap'); });
    await act(async () => { await received('proposal'); });
    expect(result.current.core.workerProposalsReady).toBe(true);
    const api = result.current.ref.current!;
    const session = result.current.core.session!;
    const layoutHere = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
    const ensureReplica = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
    try {
      const identities = await api.getParagraphIdentities();
      const paragraph = identities.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
      const initial = await api.getProposals();
      await act(async () => {
        expect(await api.proposeChanges({
          expectVersion: initial.version,
          proposals: [{
            id: 'racing-proposal', paragraph,
            suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
            op: 'insertText', at: 'start', text: 'Added text ',
          }],
        })).toMatchObject({ ok: true });
      });
      await act(async () => {
        frames.run();
        await received('sync');
      });
      const sync = posted.filter((request) => request.type === 'sync').at(-1)!;
      expect(sync).toBeDefined();
      expect(replies.has(sync.id)).toBe(true);
      const beforeToggle = await api.getProposals();
      await act(async () => {
        expect(await api.setProposalStates({
          expectVersion: beforeToggle.version,
          expectPreviewVersion: beforeToggle.previewVersion,
          changes: [{ id: 'racing-proposal', state: 'accepted' }],
        })).toMatchObject({ ok: true });
      });
      expect(session.getProposals().proposals[0]!.state).toBe('accepted');
      expect(workerProposalAuthority(session)?.holdsWorkerState()).toBe(true);
      act(() => result.current.pipeline.scheduleLayout('local'));
      await act(async () => frames.run());
      expect(result.current.errors).toEqual([]);
      expect(errorLog).not.toHaveBeenCalled();
      expect(layoutHere).not.toHaveBeenCalled();
      expect(ensureReplica).not.toHaveBeenCalled();
      expect(result.current.mainOpens).toEqual([]);
      expect(posted.filter((request) => request.type === 'sync')).toHaveLength(1);

      await act(async () => reply(sync));
      await act(async () => frames.run());
      const latest = posted.filter((request) => request.type === 'sync').at(-1)!;
      expect(latest.id).not.toBe(sync.id);
      expect(replies.has(latest.id)).toBe(true);
      await act(async () => reply(latest));
      await act(async () => {});
      await act(async () => { await frames.until(result.current.renderer.settledDisplayList(null, null)); });
      expect(isLayoutQueued(session)).toBe(false);
      expect(revisionPreviewKeyOf(result.current.pipeline.layout)).toBe(
        revisionPreviewKey(proposalRevisionPreview(session.getProposals()))
      );
      expect(sourceVersionOf(result.current.renderer.queries)).toBe(session.version());
      expect(workerProposalAuthority(session)?.geometry()?.targets['racing-proposal']).toMatchObject({ ok: true });
      expect(result.current.renderer.error).toBeNull();
      expect(result.current.errors).toEqual([]);
      expect(errorLog).not.toHaveBeenCalled();
      expect(layoutHere).not.toHaveBeenCalled();
      expect(ensureReplica).not.toHaveBeenCalled();
      expect(result.current.mainOpens).toEqual([]);
      expect(result.current.core.replicaReady).toBe(false);
      expect(workerProposalAuthority(session)?.holdsWorkerState()).toBe(true);
      expect(workers).toHaveLength(1);
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    } finally {
      layoutHere.mockRestore();
      ensureReplica.mockRestore();
    }
  } finally {
    unmount();
    errorLog.mockRestore();
    frames.restore();
  }
});

test('a lost worker during proposal sync fails the document once without rebuilding', async () => {
  const { workers, posted, received } = installWorker({ holdReply: (request) => request.type === 'sync' });
  const frames = holdFrames();
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...workerProposalProps, source: bytes },
  });
  try {
    await act(async () => { await received('open'); });
    await act(async () => { await received('bootstrap'); });
    await act(async () => { await received('proposal'); });
    const api = result.current.ref.current!;
    const session = result.current.core.session!;
    const snapshot = await api.getProposals();
    await act(async () => {
      expect(await api.setProposalStates({
        expectVersion: snapshot.version, expectPreviewVersion: snapshot.previewVersion, changes: [],
      })).toMatchObject({ ok: true });
    });
    expect(workerProposalAuthority(session)?.holdsWorkerState()).toBe(true);
    act(() => result.current.pipeline.scheduleLayout('remote'));
    await act(async () => { frames.run(); await received('sync'); });
    const waiting = result.current.renderer.settledDisplayList(null, null);
    void waiting.catch(() => {});
    await act(async () => {
      workers[0].onerror?.({ message: 'resident worker lost' } as ErrorEvent);
    });
    const failure = result.current.renderer.error;
    expect(failure?.message).toContain('resident worker lost');
    await expect(waiting).rejects.toBe(failure);
    expect(result.current.errors).toEqual([failure!]);
    expect(errorLog).toHaveBeenCalledTimes(1);
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  } finally {
    unmount();
    errorLog.mockRestore();
    frames.restore();
  }
});

test('worker-held proposals fail the document on proposal OOM without reopening source bytes', async () => {
  const options: Parameters<typeof installWorker>[0] = {};
  const { workers, posted } = installWorker(options);
  const { result, unmount } = await openWorkerProposals();
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const api = result.current.ref.current!;
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const initial = await api.getProposals();
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'held-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Held ',
        }],
      })).toMatchObject({ ok: true });
    });
    const held = await api.getProposals();
    expect(held.proposals.map(({ id }) => id)).toEqual(['held-proposal']);
    await waitFor(() => expect(
      sourceVersionOf(result.current.renderer.queries)
    ).toBe(held.version));
    options.oomStage = 'proposal';
    let failure!: ResidentWorkerOutOfMemoryError;
    await act(async () => {
      await expect(api.proposeChanges({
        expectVersion: held.version,
        proposals: [{
          id: 'next-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'end', text: ' Next',
        }],
      }).catch((error) => {
        failure = error;
        throw error;
      })).rejects.toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    });
    await expect(api.getProposals()).rejects.toBe(failure);
    await expect(api.readParagraphs({ view: 'accepted' })).rejects.toBe(failure);
    await expect(api.getProposals()).rejects.toBe(failure);
    await waitFor(() => expect(result.current.renderer.error).toBe(failure));
    expect(workers).toHaveLength(1);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(posted.filter((request) => request.type === 'proposal' &&
      request.operation.kind === 'propose')).toHaveLength(2);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(errorLog.mock.calls.filter(([, error]) => error === failure)).toEqual([
      ['[CanvasRenderer] Resident engine worker holding proposals ran out of memory', failure],
    ]);
  } finally {
    unmount();
    errorLog.mockRestore();
  }
}, 15_000);

test('a document read OOM before proposals reopens the source document once', async () => {
  const { workers, posted } = installWorker({ oomStage: 'documentRead' });
  const { result, unmount } = await openWorkerProposals();
  try {
    const initial = await result.current.ref.current!.getProposals();
    expect(initial.proposals).toEqual([]);
    const read = await result.current.ref.current!.readParagraphs({ view: 'accepted' });
    expect(read).toMatchObject({ ok: true });
    if (!read.ok) throw new Error(read.failure.message);
    expect(read.paragraphs).toHaveLength(205);
    expect(workers).toHaveLength(2);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(2);
    expect(posted.filter((request) => request.type === 'documentRead')).toHaveLength(2);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('an OOM during the first proposal starts a fresh worker instead of failing the document', async () => {
  const options: Parameters<typeof installWorker>[0] = {};
  const { workers, posted } = installWorker(options);
  const { result, unmount } = await openWorkerProposals();
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const api = result.current.ref.current!;
    const session = result.current.core.session!;
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const initial = await api.getProposals();
    options.oomStage = 'proposal';
    let outcome!: { ok: boolean } | Error;
    await act(async () => {
      outcome = await api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'first-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'First ',
        }],
      }).catch((error: Error) => error);
    });
    expect(outcome).not.toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(workers).toHaveLength(2);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(2);
    expect(result.current.renderer.error ?? null).toBeNull();
    expect(errorLog.mock.calls.some(([message]) =>
      String(message).includes('holding proposals ran out of memory')
    )).toBe(false);
    expect(registeredWorkerProposalAuthority(session)?.holdsWorkerState() ?? false).toBe(
      !(outcome instanceof Error) && outcome.ok
    );
    const after = await api.getProposals();
    expect(after.proposals.map(({ id }) => id)).toEqual(
      !(outcome instanceof Error) && outcome.ok ? ['first-proposal'] : []
    );
  } finally {
    unmount();
    errorLog.mockRestore();
  }
}, 15_000);

test('a later proposal succeeds after an OOM during the first proposal', async () => {
  const options: Parameters<typeof installWorker>[0] = {};
  const { workers, posted } = installWorker(options);
  const { result, unmount } = await openWorkerProposals();
  try {
    const api = () => result.current.ref.current!;
    const session = result.current.core.session!;
    const identities = await api().getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
    const initial = await api().getProposals();
    options.oomStage = 'proposal';
    await act(async () => {
      await expect(api().proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'failed-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Failed ',
        }],
      })).rejects.toThrow();
    });
    expect(registeredWorkerProposalAuthority(session)!.holdsWorkerState()).toBe(false);
    expect(session.getProposals().proposals).toEqual([]);
    expect(result.current.renderer.error ?? null).toBeNull();
    await act(async () => { await api().whenLayoutComplete({ timeoutMs: 5000 }); });
    const recovered = await api().getProposals();
    const recoveredParagraph = (await api().getParagraphIdentities()).paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    await act(async () => {
      expect(await api().proposeChanges({
        expectVersion: recovered.version,
        proposals: [{
          id: 'recovered-after-oom', paragraph: recoveredParagraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Recovered ',
        }],
      })).toMatchObject({ ok: true });
    });
    const proposed = await api().getProposals();
    expect(proposed.proposals.map(({ id }) => id)).toEqual(['recovered-after-oom']);
    await act(async () => {
      expect(await api().setProposalStates({
        expectVersion: proposed.version,
        expectPreviewVersion: proposed.previewVersion,
        changes: [{ id: 'recovered-after-oom', state: 'accepted' }],
      })).toMatchObject({ ok: true });
    });
    expect((await api().getProposals()).proposals).toMatchObject([
      { id: 'recovered-after-oom', state: 'accepted' },
    ]);
    expect(workers).toHaveLength(2);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(2);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.renderer.error ?? null).toBeNull();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test.each([true, false])(
  'worker revisions are asked at the replica gate, before the completion finishes, with onWorkerRevisions=%s',
  async (withCallback) => {
    let completionReplied = false;
    const asked: boolean[] = [];
    const onWorkerRevisions = mock(() => {});
    const { workers, posted } = installWorker({
      holdState: true,
      holdBootstrap: true,
      holdCompletion: true,
      revisionCount: 1,
      onRevisionCount: () => asked.push(completionReplied),
    });
    const frames = holdFrames();
    const props = {
      ...workerProposalProps,
      source: await longFixture(1200),
      followCompletion: true,
      onWorkerRevisions: withCallback ? onWorkerRevisions : undefined,
    };
    const { result, unmount } = renderHook(useHarness, { initialProps: props });
    try {
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const session = result.current.core.session!;
      const receive = workers[0].onmessage;
      workers[0].onmessage = (event) => {
        if (posted.some((request) =>
          request.type === 'completeLayout' && request.id === event.data.id
        )) completionReplied = true;
        receive?.(event);
      };
      await waitFor(() => expect(posted.map((request) => request.type)).toContain('bootstrap'));
      expect(result.current.core.workerProposalsReady).toBe(false);
      expect(posted.some((request) => request.type === 'proposal')).toBe(false);
      await act(async () => { workers[0].release(); });
      await waitFor(() => expect(posted.map((request) => request.type)).toContain('completeLayout'), {
        timeout: 5000,
      });
      await waitFor(() => expect(result.current.renderer.pendingCompletion).toBeNull());
      await act(async () => {});
      expect(posted.some((request) => request.type === 'revisionCount')).toBe(false);
      expect(onWorkerRevisions).not.toHaveBeenCalled();
      act(() => result.current.presentFrame());
      act(() => frames.run());
      act(() => frames.run());
      await waitFor(() => expect(posted.filter((request) =>
        request.type === 'revisionCount'
      )).toHaveLength(1));
      expect(asked).toEqual([false]);
      expect(completionReplied).toBe(false);
      expect(posted.findIndex((request) => request.type === 'revisionCount')).toBeGreaterThan(
        posted.findIndex((request) => request.type === 'completeLayout')
      );
      if (withCallback) {
        await waitFor(() => expect(onWorkerRevisions).toHaveBeenCalledTimes(1));
      } else {
        await waitFor(() => expect(posted.filter((request) =>
          request.type === 'encodeState'
        )).toHaveLength(1));
      }
      act(() => result.current.presentFrame());
      act(() => frames.run());
      act(() => frames.run());
      await act(async () => {});
      expect(onWorkerRevisions).toHaveBeenCalledTimes(withCallback ? 1 : 0);
      expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
      expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(
        withCallback ? 0 : 1
      );
      expect(result.current.core.replicaReady).toBe(false);
      expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
      expect(session.storyIds()).toEqual([]);
      expect(result.current.mainOpens).toEqual([]);
      await waitFor(() => expect(result.current.core.workerProposalsReady).toBe(true));
      expect(posted.filter((request) => request.type === 'proposal')).toHaveLength(1);
      expect(completionReplied).toBe(false);
      if (withCallback) {
        await act(async () => { workers[0].release(); });
        await waitFor(() => expect(result.current.core.workerProposalsReady).toBe(true));
        expect(posted.filter((request) => request.type === 'proposal')).toHaveLength(1);
        expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
        expect(result.current.core.replicaReady).toBe(false);
        expect(result.current.mainOpens).toEqual([]);
      } else {
        await act(async () => {
          workers[0].release();
          await awaitWorkerOpenReplica(session);
        });
        await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
        expect(session.hasStory('body')).toBe(true);
      }
      expect(result.current.errors).toEqual([]);
    } finally {
      unmount();
      frames.restore();
    }
  },
  15_000
);

test('worker content and preview mutations wait for the replacement frame before settling', async () => {
  const options = { holdSync: false };
  const { workers, posted } = installWorker(options);
  const { result, unmount } = await openWorkerProposals({
    ...workerProposalProps, source: await longFixture(45),
  });
  const frames = holdFrames();
  try {
    const api = result.current.ref.current!;
    const session = result.current.core.session!;
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
    const initialPages = result.current.renderer.displayList!.pages.length;
    for (const kind of ['propose', 'reject', 'withdraw'] as const) {
      const previous = result.current.renderer.displayList!;
      expect(previous.pages.some((page) => page.unbuilt)).toBe(false);
      const snapshot = await api.getProposals();
      await act(async () => {
        const changed = kind === 'propose'
          ? await api.proposeChanges({
              expectVersion: snapshot.version,
              proposals: [{
                id: 'layout-proposal', paragraph,
                suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
                op: 'insertText', at: 'start', text: 'Added '.repeat(1000),
              }],
            })
          : kind === 'reject'
            ? await api.setProposalStates({
                expectVersion: snapshot.version,
                expectPreviewVersion: snapshot.previewVersion,
                changes: [{ id: 'layout-proposal', state: 'rejected' }],
              })
            : await api.withdrawProposals({ expectVersion: snapshot.version, ids: ['layout-proposal'] });
        expect(changed).toMatchObject({ ok: true });
        if (kind === 'reject') expect(session.version()).toBe(snapshot.version);
      });
      let completed: number | null = null;
      const complete = api.whenLayoutComplete().then((pages) => { completed = pages; return pages; });
      void complete.catch(() => {});
      await act(async () => {});
      expect(completed).toBeNull();
      expect(isLayoutQueued(session)).toBe(true);
      expect(result.current.renderer.displayList).toBe(previous);

      const layouts = posted.filter((request) => request.type === 'sync').length;
      options.holdSync = true;
      await act(async () => frames.run());
      expect(posted.filter((request) => request.type === 'sync')).toHaveLength(layouts + 1);
      expect(completed).toBeNull();
      let duringLayout: number | null = null;
      const during = api.whenLayoutComplete().then((pages) => { duringLayout = pages; return pages; });
      void during.catch(() => {});
      await act(async () => {});
      expect(duringLayout).toBeNull();

      options.holdSync = false;
      await act(async () => workers[0]!.release());
      for (let frame = 0; frame < 100 && completed === null; frame += 1) {
        await act(async () => frames.run());
      }
      expect(completed).not.toBeNull();
      const pages = await complete;
      expect(await during).toBe(pages);
      expect(result.current.renderer.displayList).not.toBe(previous);
      expect(result.current.renderer.displayList!.pages.some((page) => page.unbuilt)).toBe(false);
      expect(sourceVersionOf(result.current.renderer.queries)).toBe(session.version());
      expect(isLayoutQueued(session)).toBe(false);
      if (kind === 'propose') expect(pages).toBeGreaterThan(initialPages);
      else expect(pages).toBe(initialPages);
    }
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    frames.restore();
  }
}, 15_000);

test('worker content callbacks skip preview states and refused proposals', async () => {
  const onWorkerContentChange = mock(() => {});
  const { posted } = installWorker();
  const { result, unmount } = await openWorkerProposals({
    ...workerProposalProps, onWorkerContentChange,
  });
  try {
    const api = result.current.ref.current!;
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const initial = await api.getProposals();
    expect(onWorkerContentChange).not.toHaveBeenCalled();
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'content-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Changed ',
        }],
      })).toMatchObject({ ok: true });
    });
    expect(onWorkerContentChange).toHaveBeenCalledTimes(1);
    const proposed = await api.getProposals();
    expect(proposed.version).not.toBe(initial.version);
    await act(async () => {
      expect(await api.setProposalStates({
        expectVersion: proposed.version,
        expectPreviewVersion: proposed.previewVersion,
        changes: [{ id: 'content-proposal', state: 'accepted' }],
      })).toMatchObject({ ok: true });
    });
    const preview = await api.getProposals();
    expect(preview.version).toBe(proposed.version);
    expect(preview.previewVersion).toBe(proposed.previewVersion + 1);
    expect(preview.proposals[0]!.state).toBe('accepted');
    expect(onWorkerContentChange).toHaveBeenCalledTimes(1);
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'refused-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'end', text: ' Refused',
        }],
      })).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    });
    expect(await api.getProposals()).toEqual(preview);
    expect(onWorkerContentChange).toHaveBeenCalledTimes(1);
    expect(posted.filter((request) => request.type === 'proposal' &&
      request.operation.kind === 'propose')).toHaveLength(2);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('worker proposals reach the registry before hydration and survive hand-over', async () => {
  const { posted } = installWorker();
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: workerProposalProps,
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const api = () => result.current.ref.current!;
    const identities = await api().getParagraphIdentities();
    const paragraph = identities.paragraphs.find((identity) => identity.session?.story === 'body')!.session!;
    const initial = await api().getProposals();
    let applied!: Awaited<ReturnType<DocxEditorRef['proposeChanges']>>;
    await act(async () => {
      applied = await api().proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'worker-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Proposed ',
        }],
      });
    });
    expect(applied.ok).toBe(true);
    expect(result.current.core.workerProposalsReady).toBe(true);
    expect(posted.some((request) => request.type === 'proposal' && request.operation.kind === 'propose')).toBe(true);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(session.storyIds()).toEqual([]);
    const mirrored = await api().getProposals();
    expect(mirrored.proposals.map((proposal) => proposal.id)).toEqual(['worker-proposal']);
    await act(async () => { await requestWorkerOpenReplica(session); });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.mainOpens).toEqual([false]);
    expect(session.workerDocumentMirrored()).toBe(false);
    expect(session.getProposals().proposals).toEqual(mirrored.proposals);
    const workerCalls = posted.filter((request) => request.type === 'proposal').length;
    const decided = await api().setProposalStates({
      expectVersion: session.version(),
      expectPreviewVersion: session.getProposals().previewVersion,
      changes: [{ id: 'worker-proposal', state: 'accepted' }],
    });
    expect(decided.ok).toBe(true);
    expect(session.getProposals().proposals[0]!.state).toBe('accepted');
    expect(posted.filter((request) => request.type === 'proposal')).toHaveLength(workerCalls);
  } finally {
    unmount();
    frames.restore();
  }
});

test('a sync ref call during the first in-flight proposal keeps the worker\'s proposals', async () => {
  const { workers, posted, received, reply } = installWorker({
    holdReply: (request) => request.type === 'proposal' && request.operation.kind === 'propose',
  });
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: workerProposalProps,
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const api = () => result.current.ref.current!;
    const identities = await api().getParagraphIdentities();
    const paragraph = identities.paragraphs.find((identity) => identity.session?.story === 'body')!.session!;
    const initial = await api().getProposals();
    const snapshot = await received('proposal');
    const authority = workerProposalAuthority(session)!;
    expect(authority.holdsWorkerState()).toBe(false);
    const terminate = spyOn(workers[0]!, 'terminate');
    const ensureReplica = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
    try {
      const proposed = api().proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'in-flight-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Proposed ',
        }],
      });
      void proposed.catch(() => {});
      let proposal!: ResidentEngineWorkerRequest;
      await act(async () => { proposal = await received('proposal', snapshot.id); });
      expect(proposal).toMatchObject({ operation: { kind: 'propose' } });
      act(() => { expect(() => api().getDocument()).toThrow(DocxReplicaNotReadyError); });
      await act(async () => {});
      expect(authority.holdsWorkerState()).toBe(true);
      expect(workerProposalAuthority(session)).toBe(authority);
      expect(ensureReplica).not.toHaveBeenCalled();
      expect(terminate).not.toHaveBeenCalled();
      expect(result.current.mainOpens).toEqual([]);
      expect(session.storyIds()).toEqual([]);
      expect(result.current.core.replicaReady).toBe(false);
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
      await act(async () => {
        reply(proposal);
        expect(await proposed).toMatchObject({ ok: true });
        await awaitWorkerOpenReplica(session);
      });
      await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
      expect(result.current.mainOpens).toEqual([false]);
      expect(session.workerDocumentMirrored()).toBe(false);
      const mirrored = await api().getProposals();
      expect(mirrored.proposals.map((proposal) => proposal.id)).toEqual(['in-flight-proposal']);
      await act(async () => {
        expect(await api().setProposalStates({
          expectVersion: mirrored.version,
          expectPreviewVersion: mirrored.previewVersion,
          changes: [{ id: 'in-flight-proposal', state: 'accepted' }],
        })).toMatchObject({ ok: true });
      });
      expect((await api().getProposals()).proposals[0]!.state).toBe('accepted');
      expect(result.current.errors).toEqual([]);
    } finally {
      terminate.mockRestore();
      ensureReplica.mockRestore();
    }
  } finally {
    unmount();
    frames.restore();
  }
});

test('a failed first proposal leaves the replica fallback available', async () => {
  const { workers, posted, received, reply } = installWorker({
    failProposal: true,
    holdReply: (request) => request.type === 'proposal' && request.operation.kind === 'propose',
  });
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: workerProposalProps,
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const api = () => result.current.ref.current!;
    const identities = await api().getParagraphIdentities();
    const paragraph = identities.paragraphs.find((identity) => identity.session?.story === 'body')!.session!;
    const initial = await api().getProposals();
    const snapshot = await received('proposal');
    const authority = workerProposalAuthority(session)!;
    expect(authority.holdsWorkerState()).toBe(false);
    const terminate = spyOn(workers[0]!, 'terminate');
    try {
      const proposed = api().proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'failed-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Proposed ',
        }],
      });
      void proposed.catch(() => {});
      let proposal!: ResidentEngineWorkerRequest;
      await act(async () => { proposal = await received('proposal', snapshot.id); });
      expect(proposal).toMatchObject({ operation: { kind: 'propose' } });
      expect(authority.holdsWorkerState()).toBe(true);
      expect(result.current.mainOpens).toEqual([]);
      expect(terminate).not.toHaveBeenCalled();
      await act(async () => {
        reply(proposal);
        await expect(proposed).rejects.toThrow('proposal failed');
      });
      expect(authority.holdsWorkerState()).toBe(false);
      expect(session.getProposals()).toEqual(initial);
      expect(terminate).not.toHaveBeenCalled();
      expect(result.current.core.replicaReady).toBe(false);
      act(() => { expect(api().getDocument()).not.toBeNull(); });
      expect(result.current.mainOpens).toEqual([true]);
      expect(terminate).toHaveBeenCalledTimes(1);
      expect(result.current.core.replicaReady).toBe(true);
      expect(session.workerDocumentMirrored()).toBe(false);
      expect(workerProposalAuthority(session)).toBeNull();
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
      expect((await api().getProposals()).proposals).toEqual([]);
      expect(result.current.errors).toEqual([]);
    } finally {
      terminate.mockRestore();
    }
  } finally {
    unmount();
    frames.restore();
  }
});

test('a later proposal succeeds after a terminal crash during the first proposal', async () => {
  installWorker({ crashProposalOnce: true });
  const { result, unmount } = await openWorkerProposals();
  try {
    const api = () => result.current.ref.current!;
    const session = result.current.core.session!;
    const identities = await api().getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
    const initial = await api().getProposals();
    await act(async () => {
      await expect(api().proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'failed-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Failed ',
        }],
      })).rejects.toThrow('proposal crashed');
    });
    expect(registeredWorkerProposalAuthority(session)!.holdsWorkerState()).toBe(false);
    expect(session.getProposals().proposals).toEqual([]);
    expect(result.current.renderer.error ?? null).toBeNull();
    await act(async () => { await api().whenLayoutComplete({ timeoutMs: 5000 }); });
    const recovered = await api().getProposals();
    const recoveredParagraph = (await api().getParagraphIdentities()).paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    await act(async () => {
      expect(await api().proposeChanges({
        expectVersion: recovered.version,
        proposals: [{
          id: 'recovered-after-crash', paragraph: recoveredParagraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Recovered ',
        }],
      })).toMatchObject({ ok: true });
    });
    const proposed = await api().getProposals();
    expect(proposed.proposals.map(({ id }) => id)).toEqual(['recovered-after-crash']);
    await act(async () => {
      expect(await api().setProposalStates({
        expectVersion: proposed.version,
        expectPreviewVersion: proposed.previewVersion,
        changes: [{ id: 'recovered-after-crash', state: 'accepted' }],
      })).toMatchObject({ ok: true });
    });
    expect((await api().getProposals()).proposals).toMatchObject([
      { id: 'recovered-after-crash', state: 'accepted' },
    ]);
    expect(result.current.renderer.error ?? null).toBeNull();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('a call queued behind a first proposal that crashes runs on the replacement worker', async () => {
  installWorker({ crashProposalOnce: true });
  const { result, unmount } = await openWorkerProposals();
  try {
    const api = () => result.current.ref.current!;
    const bodyParagraph = async () => (await api().getParagraphIdentities()).paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const paragraph = await bodyParagraph();
    const initial = await api().getProposals();
    const request = (id: string, expectVersion: string, target = paragraph) => ({
      expectVersion,
      proposals: [{
        id, paragraph: target,
        suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
        op: 'insertText' as const, at: 'start' as const, text: 'Queued ',
      }],
    });
    let first!: Promise<unknown>;
    let queued!: Promise<unknown>;
    await act(async () => {
      first = api().proposeChanges(request('failed-proposal', initial.version)).catch((error: Error) => error);
      queued = api().proposeChanges(request('queued-proposal', initial.version)).catch((error: Error) => error);
      expect(await first).toBeInstanceOf(Error);
    });
    let outcome: unknown;
    await act(async () => {
      outcome = await Promise.race([
        queued,
        new Promise((resolve) => setTimeout(() => resolve('hung'), 5000)),
      ]);
    });
    expect(outcome).not.toBe('hung');
    expect(outcome).toMatchObject({ ok: expect.any(Boolean) });
    await act(async () => { await api().whenLayoutComplete({ timeoutMs: 5000 }); });
    const current = await api().getProposals();
    if (!current.proposals.some(({ id }) => id === 'queued-proposal')) {
      const target = await bodyParagraph();
      await act(async () => {
        expect(await api().proposeChanges(request('queued-proposal', current.version, target)))
          .toMatchObject({ ok: true });
      });
    }
    expect((await api().getProposals()).proposals.map(({ id }) => id)).toEqual(['queued-proposal']);
    expect(result.current.renderer.error ?? null).toBeNull();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('a hand-over while a call waits for the replacement worker\'s layout settles both', async () => {
  const options: Parameters<typeof installWorker>[0] = { crashProposalOnce: true };
  const { workers, posted } = installWorker(options);
  const { result, unmount } = await openWorkerProposals();
  try {
    const api = () => result.current.ref.current!;
    const session = result.current.core.session!;
    const bodyParagraph = async () => (await api().getParagraphIdentities()).paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const initial = await api().getProposals();
    const paragraph = await bodyParagraph();
    options.holdBootstrap = true;
    let first!: Promise<unknown>;
    let queued!: Promise<unknown>;
    await act(async () => {
      first = api().proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'failed-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Failed ',
        }],
      }).catch((error: Error) => error);
      queued = api().getProposals().catch((error: Error) => error);
      expect(await first).toBeInstanceOf(Error);
    });
    await waitFor(() => expect(posted.filter((request) => request.type === 'bootstrap')).toHaveLength(2));
    expect(workers).toHaveLength(2);
    await act(async () => {});
    let replica!: Promise<unknown>;
    let outcome: unknown;
    await act(async () => {
      replica = requestWorkerOpenReplica(session)!;
      workers[1]!.release();
      outcome = await Promise.race([
        Promise.all([queued, replica]),
        new Promise((resolve) => setTimeout(() => resolve('hung'), 5000)),
      ]);
    });
    expect(outcome).not.toBe('hung');
    expect((outcome as [unknown])[0]).toMatchObject({ proposals: [] });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    const current = await api().getProposals();
    const target = await bodyParagraph();
    await act(async () => {
      expect(await api().proposeChanges({
        expectVersion: current.version,
        proposals: [{
          id: 'after-handover', paragraph: target,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Recovered ',
        }],
      })).toMatchObject({ ok: true });
    });
    expect((await api().getProposals()).proposals.map(({ id }) => id)).toEqual(['after-handover']);
    expect(result.current.renderer.error ?? null).toBeNull();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('a hand-over queued behind a recovery snapshot that fails still hydrates the calls behind it', async () => {
  let failSnapshot!: () => void;
  const { workers, posted } = installWorker({
    crashProposalOnce: true,
    failReplacementSnapshot: new Promise<void>((resolve) => { failSnapshot = resolve; }),
  });
  const { result, unmount } = await openWorkerProposals();
  try {
    const api = () => result.current.ref.current!;
    const session = result.current.core.session!;
    const bodyParagraph = async () => (await api().getParagraphIdentities()).paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const initial = await api().getProposals();
    const paragraph = await bodyParagraph();
    let first!: Promise<unknown>;
    let initializing!: Promise<unknown>;
    let queued!: Promise<unknown>;
    await act(async () => {
      first = api().proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'failed-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Failed ',
        }],
      }).catch((error: Error) => error);
      initializing = api().getProposals().catch((error: Error) => error);
      queued = api().getProposals().catch((error: Error) => error);
      expect(await first).toBeInstanceOf(Error);
    });
    await waitFor(() => expect(posted.filter((request) =>
      request.type === 'proposal' && request.operation.kind === 'snapshot'
    )).toHaveLength(2));
    expect(workers).toHaveLength(2);
    let outcome: unknown;
    await act(async () => {
      const replica = requestWorkerOpenReplica(session)!;
      failSnapshot();
      outcome = await Promise.race([
        Promise.all([initializing, queued, replica]),
        new Promise((resolve) => setTimeout(() => resolve('hung'), 5000)),
      ]);
    });
    expect(outcome).not.toBe('hung');
    const [initialized, answered] = outcome as [unknown, unknown];
    expect(initialized).toBeInstanceOf(Error);
    expect((initialized as Error).message).toBe('snapshot failed');
    expect(answered).toMatchObject({ proposals: [] });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    const current = await api().getProposals();
    const target = await bodyParagraph();
    await act(async () => {
      expect(await api().proposeChanges({
        expectVersion: current.version,
        proposals: [{
          id: 'after-handover', paragraph: target,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Recovered ',
        }],
      })).toMatchObject({ ok: true });
    });
    expect((await api().getProposals()).proposals.map(({ id }) => id)).toEqual(['after-handover']);
    expect(result.current.renderer.error ?? null).toBeNull();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
}, 15_000);

test('without worker proposals the ref waits for hydration before applying a proposal', async () => {
  const { posted, workers } = installWorker({ holdState: true });
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, workerProposals: false, allowHostProposals: true },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const session = result.current.core.session!;
  let settled = false;
  const call = result.current.ref.current!.proposeChanges({ expectVersion: session.version(), proposals: [] });
  void call.then(() => { settled = true; });
  expect(posted.some((request) => request.type === 'proposal')).toBe(false);
  expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  expect(result.current.mainOpens).toEqual([]);
  expect(settled).toBe(false);
  act(() => { requestWorkerOpenReplica(session); });
  await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
  expect(settled).toBe(false);
  await act(async () => { workers[0].release(); await call; });
  expect(settled).toBe(true);
  expect(result.current.mainOpens).toEqual([false]);
  expect(posted.some((request) => request.type === 'proposal')).toBe(false);
});

test('the host proposal gate refuses before initializing the worker authority', async () => {
  const { posted } = installWorker();
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, workerProposals: true },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const session = result.current.core.session!;
  expect(await result.current.ref.current!.proposeChanges({ expectVersion: session.version(), proposals: [] })).toEqual({
    ok: false, version: session.version(),
    failure: { code: 'read-only', message: 'The editor is read-only' },
  });
  expect(posted.some((request) => request.type === 'proposal')).toBe(false);
  expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  expect(result.current.mainOpens).toEqual([]);
});

test('a failed hand-over refuses to reseed worker proposals', async () => {
  const { workers, posted } = installWorker({ failState: true });
  const { result } = renderHook(useHarness, {
    initialProps: workerProposalProps,
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const session = result.current.core.session!;
  const api = () => result.current.ref.current!;
  const initial = await api().getProposals();
  await act(async () => {
    expect(await api().setProposalStates({
      expectVersion: initial.version, expectPreviewVersion: initial.previewVersion, changes: [],
    })).toMatchObject({ ok: true });
  });
  let failure!: Error;
  await act(async () => {
    await expect(requestWorkerOpenReplica(session)!.catch((error) => {
      failure = error;
      throw error;
    })).rejects.toThrow('open failed');
  });
  await expect(api().getProposals()).rejects.toBe(failure);
  await expect(api().readParagraphs({ view: 'accepted' })).rejects.toBe(failure);
  expect(result.current.renderer.error).toBe(failure);
  expect(workers).toHaveLength(1);
  expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
  expect(result.current.mainOpens).toEqual([]);
  expect(result.current.core.replicaReady).toBe(false);
});

test('a failed empty hand-over releases the mirror before using the main replica', async () => {
  const { received } = installWorker({ failState: true, holdReply: () => false });
  const { result } = renderHook(useHarness, {
    initialProps: workerProposalProps,
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  await act(async () => { await received('bootstrap'); });
  const session = result.current.core.session!;
  await act(async () => { await result.current.ref.current!.getProposals(); });
  expect(session.workerDocumentMirrored()).toBe(true);
  await act(async () => { await requestWorkerOpenReplica(session); });
  expect(result.current.mainOpens).toEqual([true]);
  expect(session.workerDocumentMirrored()).toBe(false);
  expect(await result.current.ref.current!.setProposalStates({
    expectVersion: session.version(), expectPreviewVersion: session.getProposals().previewVersion,
    changes: [],
  })).toMatchObject({ ok: true });
});
