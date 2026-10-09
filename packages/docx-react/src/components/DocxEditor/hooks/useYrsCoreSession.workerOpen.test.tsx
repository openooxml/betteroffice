import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createRef, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { PluginInvocation } from '../../../../../../shared/plugin-host/runtime';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import * as yrsFacade from '@betteroffice/docx/yrs';
import * as docx from '@betteroffice/docx/docx';
import type { Document } from '@betteroffice/docx/types/document';
import { createStyleResolver } from '@betteroffice/docx/styles';
import { createFontLoadScope } from '@betteroffice/docx/utils';
import {
  createYrsSession,
  preloadResidentEngineWorker,
  ResidentWorkerFailureError,
  ResidentWorkerOutOfMemoryError,
  PeerMetadataError,
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
import { createDisplayListQueries, type DisplayListQueries } from '@betteroffice/docx/layout/render';
import * as queryEngines from '@betteroffice/docx/layout/render/rustDisplayList';
import { useCanvasRenderer, type LayoutInWorker, type OpenInWorker, type OpenPreviewInWorker, type WorkerOpenedDocument } from './useDisplayList';
import type { ResolveDisplayListQueries } from './displayListQueryEpochGate';
import { useLayoutPipeline } from './useLayoutPipeline';
import { useHostSearch, type DocxSearchState } from './useHostSearch';
import { useYrsCoreSession } from './useYrsCoreSession';
import { useFileIO } from './useFileIO';
import { useDocumentLoader } from './useDocumentLoader';
import { useHistory } from '../../../hooks/useHistory';
import type { DocxEditorCollaborationOptions } from '../types';
import { awaitWorkerOpenReplica, ensureWorkerOpenReplica, requestWorkerOpenReplica, workerOpenDocumentHeld } from '../internals/workerOpenReplica';
import { DocxWorkerError } from '../internals/docxWorkerError';
import { useViewerSession } from '../internals/viewerSession';
import { isLayoutQueued, markPresented, presentedWorkerVersion, revisionPreviewKey, revisionPreviewKeyOf, sourceVersionOf, stampSourceVersion } from '../internals/layoutProvenance';
import { workerOpenSave } from '../internals/workerOpenSave';
import { workerOpenExport } from '../internals/workerOpenExport';
import * as replicaHelpers from '../internals/workerOpenReplica';
import { registeredWorkerProposalAuthority, workerProposalAuthority, workerProposalFailure } from '../internals/workerProposalAuthority';
import type { DocxEditorRef } from '../../DocxEditor';
import { PagedEditor, type PagedEditorRef } from '../PagedEditor';
import { createDocxCommandController, DocxCommandAdmissionError, UNAVAILABLE_DOCX_COMMANDS } from '../../../commands/createDocxCommandStore';
import { testBinding } from '../../../commands/testing';
import { createCommentIdAllocator } from '../commentFactories';
import { DOCX_REF_REPLICA_LOADING_MUTATIONS, DocxAsyncOnlyError, DocxReplicaNotReadyError, useDocxEditorRefApi } from './useDocxEditorRefApi';
import { usePagedEditorCommandBridge, type PagedEditorCommandBridge } from './usePagedEditorRefApi';
import { YrsInput, type YrsInputRef } from '../YrsInput';
import { flushEditorInput } from '../editorBatches';
import { defineDocxPlugin } from '../../../plugins/defineDocxPlugin';
import { createPluginClients } from '../../../plugins/createPluginClients';
import type { DocxPlugin, DocxPluginContext, DocxPluginEvent, DocxPluginSnapshot } from '../../../plugins/types';
import * as pluginHosts from '../../../plugins/useDocxPluginHost';
import * as pluginHostFactories from '../../../plugins/createDocxPluginHost';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from '../internals/engineChoice';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, renderHook, waitFor } = await import('@testing-library/react');
const { DocxEditor } = await import('../../DocxEditor');
const originalWorker = globalThis.Worker;
const originalCreateYrsSession = yrsFacade.createYrsSession;
const originalRequestAnimationFrame = globalThis.requestAnimationFrame;
const originalCancelAnimationFrame = globalThis.cancelAnimationFrame;
const originalConsoleError = console.error;
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
const globalRestores = new Set<() => void>();

function registerRestore(restore: () => void): () => void {
  const run = () => {
    if (!globalRestores.delete(run)) return;
    restore();
  };
  globalRestores.add(run);
  return run;
}

let startWorker!: () => InProcessResidentWorker;
const editModule = new WebAssembly.Module(
  new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
);
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>>;

beforeEach(() => {
  setMissingWorkerCapabilitiesForTests([]);
  const fonts = Object.getOwnPropertyDescriptor(document, 'fonts');
  registerRestore(() => {
    if (fonts) Object.defineProperty(document, 'fonts', fonts);
    else Reflect.deleteProperty(document, 'fonts');
  });
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(editModule);
});

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  startWorker = await residentWorkerFactory();
});
afterEach(() => {
  try {
    cleanup();
  } finally {
    resetEngineChoiceForTests();
    compileModule.mockRestore();
    mock.restore();
    for (const restore of [...globalRestores].reverse()) restore();
    globalThis.Worker = originalWorker;
    globalThis.requestAnimationFrame = originalRequestAnimationFrame;
    globalThis.cancelAnimationFrame = originalCancelAnimationFrame;
    console.error = originalConsoleError;
    for (const session of sessions.splice(0)) session.destroy();
  }
});
afterAll(async () => {
  await act(async () => {});
  if (ownsDom) await GlobalRegistrator.unregister();
});

function withTimeout<T>(promise: PromiseLike<T>, timeout: number, label: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`Timed out waiting for ${label} after ${timeout}ms`));
      restore();
    }, timeout);
    const restore = registerRestore(() => {
      clearTimeout(timer);
      reject(new Error(`Cancelled waiting for ${label}`));
    });
    void Promise.resolve(promise).then(
      (value) => { resolve(value); restore(); },
      (error) => { reject(error); restore(); },
    );
  });
}

function installWorker(options: {
  peerMetadata?: boolean;
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
  const postWaiters = new Set<() => void>();
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
            const request = requests.get(event.data.id);
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
        for (const waiter of postWaiters) waiter();
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
        } else send(
          request.type === 'encodeState' && !options.peerMetadata
            ? { ...request, peerMetadata: undefined }
            : request,
          transfer,
        );
      };
      workers.push(worker);
      return worker;
    }
  } as unknown as typeof Worker;
  return {
    workers, posted, hostModules, replies, responses,
    sent(type: ResidentEngineWorkerRequest['type'], afterId = 0): Promise<ResidentEngineWorkerRequest> {
      return new Promise((resolve) => {
        const check = () => {
          const request = posted.find((request) => request.type === type && request.id > afterId);
          if (!request) return;
          postWaiters.delete(check);
          resolve(request);
        };
        postWaiters.add(check);
        check();
      });
    },
    received(type: ResidentEngineWorkerRequest['type'], afterId = 0, timeout?: number): Promise<ResidentEngineWorkerRequest> {
      let check!: () => void;
      const promise = new Promise<ResidentEngineWorkerRequest>((resolve) => {
        check = () => {
          const request = [...received].find((request) => request.type === type && request.id > afterId);
          if (!request) return;
          replyWaiters.delete(check);
          resolve(request);
        };
        replyWaiters.add(check);
        check();
      });
      if (timeout === undefined) return promise;
      return withTimeout(promise, timeout, `${type} worker reply`).finally(() => {
        replyWaiters.delete(check);
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
  viewer?: boolean;
  previewFirstPage?: boolean;
  /** Opens the first-page preview in the worker, as DocxEditor does. */
  workerPreview?: boolean;
  handleSave?: () => Promise<ArrayBuffer>;
  openInWorker?: OpenInWorker;
  onWorkerOpen?: (worker: Awaited<ReturnType<OpenInWorker>>) => void;
  openPreviewInWorker?: OpenPreviewInWorker;
  source: Uint8Array;
  generation: number;
  loader?: ReturnType<typeof useDocumentLoader>;
  collaboration?: DocxEditorCollaborationOptions;
  readOnly?: boolean;
  workerProposals?: boolean;
  onWorkerRevisions?: () => void;
  onWorkerContentChange?: () => void;
  allowHostProposals?: boolean;
  resolvedCommentIds?: ReadonlySet<number>;
  measurementFont?: Uint8Array;
  styleResolver?: boolean;
  onHostDocument?: (session: YrsSession) => void;
  onLoad?: (api: DocxEditorRef) => void;
  onPresented?: (session: unknown) => void;
  /** Passes the renderer's own pending completion, as DocxEditor does. */
  followCompletion?: boolean;
  /** Holds the replica as while the shown engine's completion is still to be asked of the worker. */
  holdReplica?: boolean;
  layoutReady?: Promise<void>;
  onLayoutWait?: () => void;
  layoutCompleteSession?: YrsSession | null;
}

function useHarness(props: HarnessProps) {
  const generation = props.loader?.yrsSeedGeneration ?? props.generation;
  const relayout = useRef<(() => void) | null>(null);
  const workerRelayout = useRef<(() => void) | null>(null);
  const handoffFromRef = useRef<YrsSession | null>(null);
  const viewerSessionRef = useRef(props.viewer === true);
  viewerSessionRef.current = props.viewer === true;
  const renderer = useCanvasRenderer(
    undefined,
    props.resolvedCommentIds,
    () => relayout.current?.(),
    undefined,
    handoffFromRef,
    props.experimentalWorkerOpen,
    viewerSessionRef
  );
  useEffect(() => renderer.resetSettled(), [generation]);
  const settledDisplayList = useCallback<typeof renderer.settledDisplayList>(async (...args) => {
    props.onLayoutWait?.();
    const displayList = await renderer.settledDisplayList(...args);
    await props.layoutReady;
    if (args[3]?.aborted) throw new Error('The layout wait was cancelled');
    return displayList;
  }, [props.layoutReady, props.onLayoutWait, renderer.settledDisplayList]);
  const [host, setHost] = useState<YrsDocxHost | null>(null);
  const [, setCommentsSidebarOpen] = useState(false);
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
    const opening = (props.openInWorker ?? renderer.openInWorker)(session, source, digest, generation);
    if (!props.onWorkerOpen) return opening;
    return opening.then((worker) => {
      props.onWorkerOpen?.(worker);
      return worker;
    });
  }, [props.onWorkerOpen, props.openInWorker, renderer.openInWorker]);
  const core = useYrsCoreSession(
    true, host?.document ?? null, props.loader?.yrsSeedDocument ?? null,
    props.loader ? props.loader.yrsSeedBytes : props.source, generation, props.collaboration,
    {
      isCurrentLoad: (generation) => {
        loadChecks.current.push(generation);
        return props.loader ? props.loader.isCurrentLoad(generation) : generation === props.generation;
      },
      onSession: (session) => {
        renderer.recordSession(session);
        const open = session.openDocx.bind(session);
        session.openDocx = (input, seed, options) => {
          mainOpens.current.push(seed);
          return open(input, seed, options);
        };
      },
      onHostDocument: (host, generation, session, options) => {
        setHost(host);
        props.loader?.acceptHostDocument(host, generation, session, options);
        props.onHostDocument?.(session);
      },
      onError: (error, generation, options) => {
        notifyError(error);
        props.loader?.failHostDocument(error, generation, options);
      },
    },
    {
      previewFirstPage: props.previewFirstPage,
      heldEngine: renderer.layoutEngine,
      shownEngine: renderer.presentedEngine,
      workerOpen: props.experimentalWorkerOpen ? {
        openInWorker,
        ...(props.workerPreview ? { openPreviewInWorker: props.openPreviewInWorker ?? renderer.openPreviewInWorker } : {}),
        workerProposals: props.workerProposals,
        refreshWorkerLayout: () => workerRelayout.current?.(),
        settledDisplayList: props.layoutReady || props.onLayoutWait ? settledDisplayList : renderer.settledDisplayList,
        renderedFrame: renderer.status === 'ready' ? renderer.displayList : null,
        layoutCompleteSession: props.holdReplica ? null : props.layoutCompleteSession === undefined
          ? renderer.layoutCompleteSession : props.layoutCompleteSession,
        ...(props.holdReplica ? { pendingCompletion: renderer.presentedEngine } : {}),
        ...(props.followCompletion ? { pendingCompletion: renderer.pendingCompletion } : {}),
        viewer: props.viewer,
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
  const pipelineRef = useRef(pipeline);
  pipelineRef.current = pipeline;
  const coreRef = useRef(core);
  coreRef.current = core;
  const searchReveals = useRef<number[]>([]);
  pagedEditorRef.current = useMemo(() => core.session ? {
    getYrsSession: () => coreRef.current.session,
    getDocument: () => coreRef.current.documentFromYrs(),
    hasPendingInput: () => false,
    flushPendingInput: async () => {},
    getLayoutRequest: () => pipelineRef.current.getLayoutRequest(),
    readLayoutRequest: () => pipelineRef.current.readLayoutRequest(),
    relayout: () => pipelineRef.current.runLayoutPipeline(),
    yrsLocToDisplayPosition: (loc: Parameters<PagedEditorRef['yrsLocToDisplayPosition']>[0]) => {
      const session = coreRef.current.session!;
      const projection = createYrsPositionProjection(session, 'body');
      return yrsLocToProjectedDisplayPosition(session, () => projection, loc);
    },
    revealDisplayPosition: (position: number) => { searchReveals.current.push(position); return 'scrolled'; },
    scrollToParaId: (paraId: string) => {
      const session = coreRef.current.session!;
      if (!session.paragraphs('body').some((paragraph) => paragraph.paraId === paraId)) return false;
      const projection = createYrsPositionProjection(session, 'body');
      const position = yrsLocToProjectedDisplayPosition(session, () => projection, { story: 'body', paraId, offset: 0 });
      if (position === null) return false;
      searchReveals.current.push(position);
      return true;
    },
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
    viewerSession: props.viewer,
    ref,
    document: host?.document ?? null,
    documentFromYrs: core.documentFromYrs,
    historyStateRef: { current: host?.document ?? null },
    pagedEditorRef,
    handleSave: props.handleSave ?? (async () => new ArrayBuffer(0)),
    zoom: 1,
    setZoom: () => {},
    scrollPageInfo: { currentPage: 1, totalPages: 1, visible: true },
    loadParsedDocument: props.loader?.loadParsedDocument ?? (() => {}),
    loadBuffer: props.loader?.loadBuffer ?? (async () => {}),
    comments: [],
    setComments: () => {},
    setShowCommentsSidebar: () => {},
    contentChangeSubscribersRef: { current: new Set() },
    selectionChangeSubscribersRef: { current: new Set() },
    getCachedStyleResolver: props.styleResolver ? createStyleResolver : (() => { throw new Error('unused'); }) as never,
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
    viewerSession: props.viewer,
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

async function peerHydrationSource(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await longFixture(1));
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>',
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
  zip.file('word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${ns}><w:body><w:p w14:paraId="0000B001">` +
    '<w:commentRangeStart w:id="1"/><w:r><w:t>Peer paragraph</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
    '<w:r><w:commentReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>');
  zip.file('word/comments.xml', `<w:comments ${ns}><w:comment w:id="1" w:author="Peer author" w:date="2026-10-01T00:00:00Z">` +
    '<w:p w14:paraId="0000C001"><w:r><w:t>Peer comment</w:t></w:r></w:p></w:comment></w:comments>');
  return zip.generateAsync({ type: 'uint8array' });
}

async function peerHydrationHarness(metadata: boolean, sourceOverride?: Uint8Array) {
  const worker = installWorker({ peerMetadata: metadata, holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const source = sourceOverride ?? await peerHydrationSource();
  let snapshot: Awaited<ReturnType<NonNullable<WorkerOpenedDocument['encodeVersionedState']>>> | undefined;
  let corruptOffset: number | undefined;
  let fallback: ReturnType<typeof mock> | undefined;
  const replicas: Array<YrsSession | null> = [];
  const props: HarnessProps = {
    ...initialProps, source,
    collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    onWorkerOpen: (opened) => {
      if (!opened?.encodeVersionedState) throw new Error('Expected a versioned worker snapshot');
      const encode = opened.encodeVersionedState;
      opened.encodeVersionedState = async (prefetch) => {
        snapshot = await encode(prefetch);
        if (corruptOffset !== undefined) {
          if (!snapshot.metadata) throw new Error('No metadata to corrupt');
          snapshot = { ...snapshot, metadata: snapshot.metadata.slice() };
          snapshot.metadata![corruptOffset] ^= 1;
        }
        return snapshot;
      };
      const openedFallback = spyOn(opened, 'fallback');
      fallback ??= openedFallback;
      registerRestore(() => openedFallback.mockRestore());
    },
  };
  const view = renderHook(useHarness, { initialProps: props });
  await waitFor(() => expect(view.result.current.host).not.toBeNull());
  const session = view.result.current.core.session!;
  const bootstrap = spyOn(session, 'bootstrapPeer');
  const load = spyOn(session, 'loadState');
  const open = spyOn(session, 'openDocx');
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  registerRestore(() => { bootstrap.mockRestore(); load.mockRestore(); open.mockRestore(); warn.mockRestore(); });
  return {
    ...view, worker, frames, tasks, props, source, session, bootstrap, load, open, warn, replicas,
    get snapshot() {
      if (!snapshot) throw new Error('No captured peer snapshot');
      return snapshot;
    },
    get fallback() {
      if (!fallback) throw new Error('No opened worker');
      return fallback;
    },
    corruptMetadata(offset: number) { corruptOffset = offset; },
    async start() {
      const pending = requestWorkerOpenReplica(session);
      if (!pending) throw new Error('No pending peer hydration');
      void pending.catch(() => {});
      await waitFor(() => expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(true));
      await act(async () => {
        await withTimeout(Promise.resolve(worker.workers[0]!.release()), 1000, 'held peer snapshot release');
      });
      await waitFor(() => expect(tasks.tasks).toHaveLength(1));
      return { pending };
    },
    async finish(pending: Promise<void>) {
      await act(async () => tasks.run());
      await waitFor(() => expect(tasks.tasks).toHaveLength(1));
      await act(async () => {
        replicaHelpers.notifyWorkerOpenLayoutProgress(session, 'complete');
        await tasks.run();
        await withTimeout(pending, 1000, 'peer hydration readiness');
      });
    },
    close() {
      view.unmount();
      visibility.restore();
      tasks.restore();
      frames.restore();
    },
  };
}

test('metadata hydration bootstraps the peer and matches compatibility paragraphs, ids and comment exports', async () => {
  const peer = await peerHydrationHarness(true);
  try {
    const { pending } = await peer.start();
    expect(peer.snapshot.metadata).toBeDefined();
    const snapshot = peer.snapshot;
    const host = peer.result.current.host;
    if (snapshot.metadata === undefined || host === null) throw new Error('Expected peer metadata and host');
    expect(peer.bootstrap).toHaveBeenCalledTimes(1);
    expect(peer.bootstrap.mock.calls[0]![0]).toBe(snapshot.state);
    expect(peer.bootstrap.mock.calls[0]![1]).toBe(snapshot.metadata);
    expect(peer.bootstrap.mock.calls[0]![2]).toEqual(peer.source);
    expect(peer.bootstrap.mock.calls[0]![3]).toBe(host);
    expect(peer.open).not.toHaveBeenCalled();
    expect(peer.load).not.toHaveBeenCalled();
    expect(peer.result.current.core.replicaReady).toBe(false);
    await peer.finish(pending);
    expect(peer.result.current.core.replicaReady).toBe(true);
    expect(peer.replicas).toEqual([peer.session]);
    expect(peer.load).not.toHaveBeenCalled();
    expect(peer.warn).not.toHaveBeenCalled();
    expect(peer.fallback).not.toHaveBeenCalled();
    const baseline = await createYrsSession({ clientId: peer.session.clientId });
    sessions.push(baseline);
    baseline.openDocx(peer.source, false);
    baseline.loadState(peer.snapshot.state);
    expect(peer.session.storyIds()).toEqual(baseline.storyIds());
    for (const story of baseline.storyIds()) {
      expect(peer.session.paragraphs(story)).toEqual(baseline.paragraphs(story));
      expect(peer.session.storyParagraphIds(story)).toEqual(baseline.storyParagraphIds(story));
    }
    expect(peer.session.paragraphIdentities()).toEqual(baseline.paragraphIdentities());
    expect(peer.session.listComments()).toEqual(baseline.listComments());
    const options = { revisionView: 'markup' as const, stories: ['comments' as const] };
    const actual = peer.session.exportStructured(options);
    const expected = baseline.exportStructured(options);
    if (!actual.ok || !expected.ok) throw new Error('Comment export failed');
    expect(actual.content.stories).toHaveLength(1);
    expect(actual.content).toEqual(expected.content);
    expect(peer.result.current.errors).toEqual([]);
  } finally {
    peer.close();
  }
});

test('metadata absence uses the captured state and warns exactly once with the reason and tags', async () => {
  const peer = await peerHydrationHarness(false);
  try {
    const { pending } = await peer.start();
    expect(peer.bootstrap).not.toHaveBeenCalled();
    expect(peer.open).toHaveBeenCalledTimes(1);
    expect(peer.open).toHaveBeenCalledWith(peer.source, false);
    expect(peer.load).not.toHaveBeenCalled();
    await peer.finish(pending);
    expect(peer.load).toHaveBeenCalledTimes(1);
    expect(peer.load.mock.calls[0]![0]).toBe(peer.snapshot.state);
    expect(peer.warn).toHaveBeenCalledTimes(1);
    expect(peer.warn.mock.calls[0]![0]).toContain('missing-capability: Worker omitted peer metadata');
    expect(peer.warn.mock.calls[0]![0]).toContain('expected tag v1/');
    expect(peer.warn.mock.calls[0]![0]).toContain('received tag absent');
    expect(peer.fallback).not.toHaveBeenCalled();
    expect(peer.result.current.core.replicaReady).toBe(true);
  } finally {
    peer.close();
  }
});

test.each(['bootstrap-rejection', 'unsupported-version', 'shape-mismatch'] as const)(
  'metadata %s falls back on the same session and state without worker fallback', async (reason) => {
    const peer = await peerHydrationHarness(true);
    try {
      if (reason === 'bootstrap-rejection') {
        peer.bootstrap.mockImplementation(() => { throw new PeerMetadataError('source-mismatch', 'Rejected source'); });
      } else {
        peer.corruptMetadata(reason === 'unsupported-version' ? 8 : 12);
      }
      const { pending } = await peer.start();
      expect(peer.bootstrap).toHaveBeenCalledTimes(1);
      expect(peer.open).toHaveBeenCalledWith(peer.source, false);
      expect(peer.load).not.toHaveBeenCalled();
      await peer.finish(pending);
      expect(peer.load).toHaveBeenCalledTimes(1);
      expect(peer.load.mock.calls[0]![0]).toBe(peer.snapshot.state);
      expect(peer.result.current.core.session).toBe(peer.session);
      expect(peer.open.mock.calls.map((call) => call[1])).toEqual([false]);
      expect(peer.fallback).not.toHaveBeenCalled();
      expect(peer.warn).toHaveBeenCalledTimes(1);
      expect(peer.warn.mock.calls[0]![0]).toContain(reason === 'bootstrap-rejection' ? 'source-mismatch' : reason);
      expect(peer.warn.mock.calls[0]![0]).toContain('expected tag v1/');
      expect(peer.warn.mock.calls[0]![0]).toContain(`received tag v${reason === 'unsupported-version' ? 0 : 1}/`);
      expect(peer.result.current.errors).toEqual([]);
      expect(peer.result.current.core.replicaReady).toBe(true);
    } finally {
      peer.close();
    }
  },
);

test('invalid table metadata falls back on the same session and state with one warning', async () => {
  const peer = await peerHydrationHarness(true);
  try {
    peer.bootstrap.mockImplementation(() => {
      throw new PeerMetadataError('invalid-metadata', 'invalid table cell column');
    });
    const { pending } = await peer.start();
    expect(peer.bootstrap).toHaveBeenCalledTimes(1);
    expect(peer.open).toHaveBeenCalledWith(peer.source, false);
    expect(peer.load).not.toHaveBeenCalled();
    await peer.finish(pending);
    expect(peer.load).toHaveBeenCalledTimes(1);
    expect(peer.load.mock.calls[0]![0]).toBe(peer.snapshot.state);
    expect(peer.result.current.core.session).toBe(peer.session);
    expect(peer.open.mock.calls.map((call) => call[1])).toEqual([false]);
    expect(peer.fallback).not.toHaveBeenCalled();
    expect(peer.warn).toHaveBeenCalledTimes(1);
    expect(peer.warn.mock.calls[0]![0]).toContain('invalid-metadata: invalid table cell column');
    expect(peer.warn.mock.calls[0]![0]).toContain('expected tag v1/');
    expect(peer.warn.mock.calls[0]![0]).toContain('received tag v1/');
    expect(peer.result.current.errors).toEqual([]);
    expect(peer.result.current.core.replicaReady).toBe(true);
  } finally {
    peer.close();
  }
});

test.each(['replace', 'unmount'] as const)('a document %s during metadata bootstrap discards hydration', async (action) => {
  const peer = await peerHydrationHarness(true);
  try {
    const { pending } = await peer.start();
    expect(peer.bootstrap).toHaveBeenCalledTimes(1);
    expect(peer.result.current.core.replicaReady).toBe(false);
    expect(peer.replicas).toEqual([]);
    const errors = peer.result.current.errors;
    if (action === 'replace') {
      act(() => peer.rerender({ ...peer.props, source: peer.source.slice(), generation: 2 }));
      await waitFor(() => expect(peer.result.current.core.sessionGeneration).toBe(2));
    } else peer.unmount();
    await act(async () => peer.tasks.run());
    await expect(withTimeout(pending, 1000, 'cancelled peer hydration')).rejects.toThrow('The document changed');
    expect(replicaHelpers.workerOpenReplicaPending(peer.session)).toBe(false);
    expect(peer.load).not.toHaveBeenCalled();
    expect(peer.open).not.toHaveBeenCalled();
    expect(peer.warn).not.toHaveBeenCalled();
    expect(peer.fallback).not.toHaveBeenCalled();
    expect(peer.replicas).toEqual([]);
    expect(errors).toEqual([]);
    if (action === 'replace') expect(peer.result.current.core.replicaReady).toBe(false);
  } finally {
    peer.close();
  }
});

test.each([
  ['fractional omissions and span', '0.5', '1.5', '2.5', 0, 1, 2],
  ['fractional span truncates to one', '1.5', '0.5', '1.5', 1, 0, 1],
  ['span above u32 clamps to u16', '0.5', '0.5', '4294967296.5', 0, 0, 65535],
] as const)(
  'parser-accepted comment table bootstraps without fallback and exports like source open: %s',
  async (_name, before, after, span, gridBefore, gridAfter, gridSpan) => {
    const zip = await JSZip.loadAsync(await peerHydrationSource());
    const comments = await zip.file('word/comments.xml')!.async('string');
    const table = '<w:tbl><w:tr><w:trPr>' +
      `<w:gridBefore w:val="${before}"/><w:gridAfter w:val="${after}"/>` +
      '</w:trPr><w:tc><w:tcPr>' +
      `<w:gridSpan w:val="${span}"/>` +
      '</w:tcPr><w:p><w:r><w:t>Normalized comment cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    zip.file('word/comments.xml', comments.replace('</w:comment>', table + '</w:comment>'));
    const source = await zip.generateAsync({ type: 'uint8array' });
    const peer = await peerHydrationHarness(true, source);
    try {
      const { pending } = await peer.start();
      expect(peer.snapshot.metadata).toBeDefined();
      expect(peer.bootstrap).toHaveBeenCalledTimes(1);
      expect(peer.open).not.toHaveBeenCalled();
      expect(peer.load).not.toHaveBeenCalled();
      await peer.finish(pending);
      expect(peer.warn).not.toHaveBeenCalled();
      expect(peer.fallback).not.toHaveBeenCalled();
      expect(peer.result.current.core.replicaReady).toBe(true);
      const baseline = await createYrsSession({ clientId: peer.session.clientId });
      sessions.push(baseline);
      baseline.openDocx(source, false);
      baseline.loadState(peer.snapshot.state);
      for (const revisionView of ['accepted', 'original', 'markup'] as const) {
        const options = { revisionView, stories: ['comments' as const] };
        const actual = peer.session.exportStructured(options);
        const expected = baseline.exportStructured(options);
        if (!actual.ok || !expected.ok) throw new Error('Comment table export failed');
        expect(actual.content).toEqual(expected.content);
        const block = actual.content.stories[0]!.blocks.find((block) => block.kind === 'table');
        if (!block || block.kind !== 'table') throw new Error('Expected a comment table');
        expect(block.table.gridColumns).toBe(gridBefore + gridSpan + gridAfter);
        expect(block.table.rows[0]).toMatchObject({ gridBefore, gridAfter });
        expect(block.table.rows[0]!.cells[0]).toMatchObject({ column: gridBefore, gridSpan });
      }
      expect(peer.result.current.errors).toEqual([]);
    } finally {
      peer.close();
    }
  },
);

function observeMetadataBootstrap(session: YrsSession) {
  const bootstrap = spyOn(session, 'bootstrapPeer');
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  registerRestore(() => { bootstrap.mockRestore(); warn.mockRestore(); });
  return {
    bootstrap,
    check(installed = true) {
      expect(bootstrap).toHaveBeenCalledTimes(installed ? 1 : 0);
      expect(warn.mock.calls.filter(([message]) =>
        String(message).includes('peer hydration compatibility mode')
      )).toEqual([]);
    },
  };
}

test.each([1, 2])('synchronous ensure finishes the worker peer at hydration yield %s once with metadata bootstrap', async (boundary) => {
  const { workers, posted } = installWorker({ peerMetadata: true, holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  try {
    const { result } = renderHook(useHarness, { initialProps: {
      ...initialProps,
      collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    } });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const metadata = observeMetadataBootstrap(session);
    const load = spyOn(session, 'loadState');
    restoreLoad = registerRestore(() => load.mockRestore());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    if (boundary === 2) await act(async () => tasks.run());
    act(() => {
      ensureWorkerOpenReplica(session);
      expect(session.hasStory('body')).toBe(true);
      expect(result.current.core.replicaReadyRef?.current).toBe(true);
      expect(load).not.toHaveBeenCalled();
      metadata.check();
    });
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.mainOpens).toEqual([]);
    expect(replicas).toEqual([session]);
    await act(async () => {
      await tasks.run();
      await pending;
    });
    expect(session.hasStory('body')).toBe(true);
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.core.replicaReadyRef?.current).toBe(true);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([]);
    expect(replicas).toEqual([session]);
    metadata.check();
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test.each([false, true])('metadata bootstrap bypasses a loadState error after yielding with failure=%s', async (fails) => {
  const { workers, posted } = installWorker({ peerMetadata: true, holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  let restoreOpen = () => {};
  try {
    const { result } = renderHook(useHarness, { initialProps: {
      ...initialProps,
      collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    } });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const metadata = observeMetadataBootstrap(session);
    const loadError = new Error('State load failed');
    const fallbackError = new Error('Replica fallback failed');
    const load = spyOn(session, 'loadState').mockImplementation(() => { throw loadError; });
    restoreLoad = registerRestore(() => load.mockRestore());
    const originalOpen = session.openDocx.bind(session);
    const fallbackSeeds: boolean[] = [];
    const open = spyOn(session, 'openDocx').mockImplementation((source, seed, options) => {
      fallbackSeeds.push(seed);
      if (fails && seed) throw fallbackError;
      return originalOpen(source, seed, options);
    });
    restoreOpen = registerRestore(() => open.mockRestore());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    expect(load).not.toHaveBeenCalled();
    metadata.check();
    expect(result.current.errors).toEqual([]);
    await act(async () => {
      await tasks.run();
      await tasks.run();
      replicaHelpers.notifyWorkerOpenLayoutProgress(session, 'complete');
      await pending;
    });
    expect(load).not.toHaveBeenCalled();
    expect(fallbackSeeds).toEqual([]);
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.core.replicaReadyRef?.current).toBe(true);
    expect(replicas).toEqual([session]);
    metadata.check();
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      restoreOpen();
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test.each(['replace', 'unmount'] as const)('a document %s between hydration tasks stops the stale peer with metadata bootstrap', async (action) => {
  const { workers, posted } = installWorker({ peerMetadata: true, holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  try {
    const props = { ...initialProps,
      collaboration: { onReplica: (session: unknown) => replicas.push(session as YrsSession | null) },
    };
    const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const metadata = observeMetadataBootstrap(session);
    const errors = result.current.errors;
    const load = spyOn(session, 'loadState');
    restoreLoad = registerRestore(() => load.mockRestore());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    expect(result.current.mainOpens).toEqual([]);
    metadata.check();
    expect(session.hasStory('body')).toBe(true);
    if (action === 'replace') {
      rerender({ ...props, source: bytes.slice(), generation: 2 });
      await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
      expect(result.current.core.session).not.toBe(session);
    } else {
      unmount();
    }
    await act(async () => tasks.run());
    await expect(pending).rejects.toThrow('The document changed');
    expect(load).not.toHaveBeenCalled();
    metadata.check();
    expect(replicas).toEqual([]);
    expect(errors).toEqual([]);
    if (action === 'replace') {
      expect(result.current.mainOpens).toEqual([]);
      expect(result.current.core.replicaReady).toBe(false);
      expect(result.current.core.replicaReadyRef?.current).toBe(false);
    }
  } finally {
    try {
      cleanup();
    } finally {
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test('the editor state prefetch waits for a sliced background page-build reply with metadata bootstrap', async () => {
  const options = {
    peerMetadata: true,
    holdCompletion: true,
    holdReply: (request: ResidentEngineWorkerRequest) =>
      request.type === 'buildPages' || request.type === 'encodeState',
  };
  const { posted, workers, received, reply, replies } = installWorker(options);
  const frames = holdFrames(true);
  const fallback = holdPeerFallback();
  const visibility = stubDocumentVisibility('visible');
  try {
    const props = { ...initialProps, source: await longFixture(1200) };
    const { result, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const metadata = observeMetadataBootstrap(session);
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    await waitFor(() => expect(posted.some((request) => request.type === 'completeLayout')).toBe(true));
    act(() => result.current.presentFrame());
    const layout = result.current.renderer.settledDisplayList(null, null, 'document');
    void layout.catch(() => {});
    options.holdCompletion = false;
    await act(async () => workers[0].release());
    let batch = await received('buildPages');
    while (batch.type === 'buildPages' && !batch.background) {
      await act(async () => reply(batch));
      batch = await received('buildPages', batch.id);
    }
    if (batch.type !== 'buildPages') throw new Error('expected a page build');
    expect(batch.background).toBe(true);
    expect(batch.pages.length).toBeGreaterThan(4);
    expect(batch.pages).toEqual(expect.arrayContaining([5, 6]));
    expect(replies.has(batch.id)).toBe(true);
    const window = result.current.renderer.settledDisplayList(null, null, 'window');
    void window.catch(() => {});
    await act(async () => {
      frames.run();
      frames.runIdle();
    });
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    const postedBeforeReply = posted.length;
    await act(async () => reply(batch));
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    const prefetched = await received('encodeState');
    expect(posted.indexOf(prefetched)).toBeGreaterThanOrEqual(postedBeforeReply);
    expectStatePrefetchAfterPageBuilds(posted);
    await frames.untilCommitted(layout);
    await window;
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    await act(async () => {
      reply(prefetched);
      frames.runIdle();
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expectStatePrefetchAfterPageBuilds(posted);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([]);
    metadata.check();
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    visibility.restore();
    fallback.restore();
    frames.restore();
  }
}, 15_000);

test('the editor peer waits for window layout to settle and then starts on idle with metadata bootstrap', async () => {
  let holdMargin = true;
  let marginId: number | null = null;
  const options = {
    peerMetadata: true,
    holdCompletion: true,
    holdReply: (request: ResidentEngineWorkerRequest) => {
      if (!holdMargin || request.type !== 'buildPages' || !request.background) return false;
      marginId ??= request.id;
      return request.id === marginId;
    },
  };
  const { posted, workers, received, reply } = installWorker(options);
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  try {
    const props = { ...initialProps, source: await longFixture(1200) };
    const { result, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const full = result.current.core.session!;
    const metadata = observeMetadataBootstrap(full);
    const load = spyOn(full, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
    await waitFor(() => expect(posted.some((request) => request.type === 'completeLayout')).toBe(true));
    act(() => result.current.presentFrame());
    let settled = false;
    const layout = result.current.renderer.settledDisplayList(null, null, 'window');
    void layout.then(() => { settled = true; }, () => {});
    await act(async () => {
      frames.run();
      frames.runIdle();
    });
    expect(settled).toBe(false);
    expect([...frames.idleCallbacks.values()].some(({ options }) => options?.timeout === 2000)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.core.replicaReady).toBe(false);
    options.holdCompletion = false;
    await act(async () => workers[0].release());
    await waitFor(() => expect(posted.some((request) => request.type === 'buildPages')).toBe(true));
    await waitFor(() => expect(marginId).not.toBeNull());
    const margin = await received('buildPages', marginId! - 1);
    expect(margin).toMatchObject({ type: 'buildPages', pages: [5, 6], background: true });
    expect(settled).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    holdMargin = false;
    await act(async () => reply(margin));
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(posted.findIndex((request) => request.type === 'encodeState')).toBeGreaterThan(posted.indexOf(margin));
    await frames.untilCommitted(layout);
    expect(settled).toBe(true);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect([...frames.idleCallbacks.values()].some(({ options }) => options?.timeout === 2000)).toBe(true);
    await act(async () => frames.runIdle());
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(posted.some((request) => request.type === 'encodeState')).toBe(true);
    expect(result.current.mainOpens).toEqual([]);
    metadata.check();
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    visibility.restore();
    frames.restore();
  }
}, 15_000);

test.each(['resolved', 'in flight'] as const)(
  'a host change after state prefetch is %s hydrates the current worker state with metadata bootstrap',
  async (stage) => {
    const worker = installWorker({ peerMetadata: true, holdReply: (request) => request.type === 'encodeState' });
    const frames = holdFrames(true);
    const visibility = stubDocumentVisibility('visible');
    let resident!: NonNullable<Awaited<ReturnType<OpenInWorker>>>;
    const openInWorker: OpenInWorker = async (...args) => {
      const opened = await result.current.renderer.openInWorker(...args);
      if (opened) resident = opened;
      return opened;
    };
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, source: await longFixture(1), openInWorker },
    });
    try {
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const session = result.current.core.session!;
      const metadata = observeMetadataBootstrap(session);
      const load = spyOn(session, 'loadState');
      registerRestore(() => load.mockRestore());
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
      act(() => result.current.presentFrame());
      await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
      const prefetched = await worker.received('encodeState');
      const before = worker.responses.get(prefetched)!;
      if (!before.ok || !before.state || !before.peerMetadata) throw new Error('expected prefetched state');
      expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
      expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
      expect(result.current.mainOpens).toEqual([]);
      expect(load).not.toHaveBeenCalled();
      if (stage === 'resolved') await act(async () => worker.reply(prefetched));
      const identities = await resident.documentRead({ kind: 'paragraphIdentities' });
      const paragraph = identities.value.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
      await act(async () => {
        const changed = await resident.proposal({
          kind: 'propose',
          request: {
            expectVersion: identities.version,
            proposals: [{
              id: 'after-prefetch', paragraph,
              suggest: { author: 'Host', date: '2026-10-04T00:00:00Z' },
              op: 'insertText', at: 'start', text: 'Changed ',
            }],
          },
        });
        expect(changed.result).toMatchObject({ ok: true });
      });
      expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
      await act(async () => frames.runIdle());
      await waitFor(() => expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(2));
      const fresh = await worker.received('encodeState', prefetched.id);
      const after = worker.responses.get(fresh)!;
      if (!after.ok || !after.state || !after.peerMetadata) throw new Error('expected current worker state');
      expect(after.version).not.toBe(before.version);
      expect(load).not.toHaveBeenCalled();
      await act(async () => {
        if (stage === 'in flight') worker.reply(prefetched);
        worker.reply(fresh);
        await awaitWorkerOpenReplica(session);
      });
      expect(load).not.toHaveBeenCalled();
      expect(metadata.bootstrap.mock.calls[0]![0]).toEqual(new Uint8Array(after.state));
      expect(metadata.bootstrap.mock.calls[0]![0]).not.toEqual(new Uint8Array(before.state));
      expect(metadata.bootstrap.mock.calls[0]![1]).toEqual(new Uint8Array(after.peerMetadata));
      expect(metadata.bootstrap.mock.calls[0]![1].buffer).toBe(after.peerMetadata);
      expect(metadata.bootstrap.mock.calls[0]![1].buffer).not.toBe(before.peerMetadata);
      expect(session.paragraphs('body')[0].text).toBe('Changed First paragraph');
      expect(result.current.mainOpens).toEqual([]);
      expect(result.current.core.replicaReady).toBe(true);
      metadata.check();
      expect(result.current.errors).toEqual([]);
    } finally {
      unmount();
      visibility.restore();
      frames.restore();
    }
  },
  15_000
);

test('an already settled editor window prefetches state when its peer begins waiting with metadata bootstrap', async () => {
  const { posted, received, reply } = installWorker({ peerMetadata: true, holdReply: (request) => request.type === 'encodeState' });
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const onLayoutWait = mock(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, source: await longFixture(1), onLayoutWait },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const metadata = observeMetadataBootstrap(session);
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    act(() => result.current.presentFrame());
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
    const prefetched = await received('encodeState');
    expect(onLayoutWait).toHaveBeenCalledTimes(1);
    expect(posted.some((request) => request.type === 'buildPages')).toBe(false);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    await act(async () => reply(prefetched));
    expect(load).not.toHaveBeenCalled();
    await act(async () => frames.runIdle());
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([]);
    metadata.check();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    visibility.restore();
    frames.restore();
  }
});

test('a failed state prefetch takes the encode-failure fallback without encoding again with metadata bootstrap', async () => {
  let holdStateReply = true;
  const { workers, posted, received, reply, replies } = installWorker({ peerMetadata: true,
    holdReply: (request) => holdStateReply && request.type === 'encodeState',
  });
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const { result, rerender, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, source: await longFixture(1) },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const metadata = observeMetadataBootstrap(session);
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    act(() => result.current.presentFrame());
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
    const prefetched = await received('encodeState');
    const lateState = replies.get(prefetched.id)!;
    await act(async () => {
      workers[0].onmessage?.({
        data: { id: prefetched.id, ok: false, error: 'prefetch failed' },
      } as MessageEvent);
      reply(prefetched);
    });
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.errors).toEqual([]);
    await act(async () => frames.runIdle());
    await waitFor(() => expect(result.current.mainOpens).toEqual([true]));
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.errors).toEqual([]);
    await act(async () => lateState());
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([true]);
    metadata.check(false);
    holdStateReply = false;
    act(() => rerender({ ...initialProps, source: longBytes.slice(), generation: 2 }));
    await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const replacement = result.current.core.session!;
    const replacementMetadata = observeMetadataBootstrap(replacement);
    const pending = requestWorkerOpenReplica(replacement)!;
    await act(async () => {
      replicaHelpers.notifyWorkerOpenLayoutProgress(replacement, 'complete');
      await pending;
    });
    replacementMetadata.check();
    expect(result.current.core.replicaReady).toBe(true);
  } finally {
    unmount();
    visibility.restore();
    frames.restore();
  }
});

test.each([
  ['replace', 'resolved'], ['replace', 'in flight'],
  ['unmount', 'resolved'], ['unmount', 'in flight'],
] as const)(
  'a document %s discards a %s state prefetch before peer start with metadata bootstrap',
  async (action, stage) => {
    let holdStateReply = true;
    const { posted, received, reply, responses } = installWorker({ peerMetadata: true,
      holdReply: (request) => holdStateReply && request.type === 'encodeState',
    });
    const frames = holdFrames(true);
    const visibility = stubDocumentVisibility('visible');
    const { result, rerender, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, source: await longFixture(1) },
    });
    try {
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const previous = result.current.core.session!;
      const previousMetadata = observeMetadataBootstrap(previous);
      const loadPrevious = spyOn(previous, 'loadState');
      registerRestore(() => loadPrevious.mockRestore());
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(previous));
      act(() => result.current.presentFrame());
      await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
      const prefetched = await received('encodeState');
      const oldState = responses.get(prefetched)!;
      if (!oldState.ok || !oldState.state || !oldState.peerMetadata) throw new Error('expected previous document state');
      if (stage === 'resolved') await act(async () => reply(prefetched));
      expect(replicaHelpers.workerOpenReplicaStarted(previous)).toBe(false);
      const staleIdle = [...frames.idleCallbacks.values()].filter(({ options }) => options?.timeout === 2000);
      expect(staleIdle).toHaveLength(1);
      holdStateReply = false;
      if (action === 'replace') {
        const source = await longFixture(2);
        act(() => rerender({ ...initialProps, source, generation: 2 }));
        await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
      } else unmount();
      await act(async () => {
        if (stage === 'in flight') reply(prefetched);
        for (const { callback } of staleIdle) callback({ didTimeout: false, timeRemaining: () => 50 });
        frames.runIdle();
      });
      expect(loadPrevious).not.toHaveBeenCalled();
      expect(replicaHelpers.workerOpenReplicaPending(previous)).toBe(false);
      expect(result.current.mainOpens).toEqual([]);
      if (action === 'replace') {
        const replacement = result.current.core.session!;
        const metadata = observeMetadataBootstrap(replacement);
        const load = spyOn(replacement, 'loadState');
        registerRestore(() => load.mockRestore());
        act(() => result.current.pipeline.runLayoutPipeline());
        await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(replacement));
        act(() => result.current.presentFrame());
        await frames.settleAndIdle(result.current.renderer.settledDisplayList(null, null, 'window'));
        await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
        expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(2);
        expect(load).not.toHaveBeenCalled();
        expect(metadata.bootstrap.mock.calls[0]![0]).not.toEqual(new Uint8Array(oldState.state));
        metadata.check();
        expect(texts(replacement).body).toEqual(['First paragraph', 'Tail paragraph']);
        expect(loadPrevious).not.toHaveBeenCalled();
        expect(result.current.mainOpens).toEqual([]);
      } else expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
      previousMetadata.check(false);
      expect(result.current.errors).toEqual([]);
    } finally {
      if (action === 'replace') unmount();
      visibility.restore();
      frames.restore();
    }
  },
  15_000
);

test('worker proposals reach the registry before hydration and survive hand-over with metadata bootstrap', async () => {
  const { posted } = installWorker({ peerMetadata: true });
  const frames = holdFrames();
  const { result, rerender, unmount } = renderHook(useHarness, {
    initialProps: workerProposalProps,
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const metadata = observeMetadataBootstrap(session);
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
    act(() => rerender({ ...workerProposalProps, viewer: false }));
    await act(async () => { await requestWorkerOpenReplica(session); });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.mainOpens).toEqual([]);
    expect(session.workerDocumentMirrored()).toBe(false);
    expect(session.getProposals().proposals).toEqual(mirrored.proposals);
    expect((await api().getProposals()).proposals).toEqual(mirrored.proposals);
    const workerCalls = posted.filter((request) => request.type === 'proposal').length;
    const decided = await api().setProposalStates({
      expectVersion: session.version(),
      expectPreviewVersion: session.getProposals().previewVersion,
      changes: [{ id: 'worker-proposal', state: 'accepted' }],
    });
    expect(decided.ok).toBe(true);
    expect(session.getProposals().proposals[0]!.state).toBe('accepted');
    expect((await api().getProposals()).proposals[0]!.state).toBe('accepted');
    expect(posted.filter((request) => request.type === 'proposal')).toHaveLength(workerCalls);
    metadata.check();
  } finally {
    unmount();
    frames.restore();
  }
});

async function openingPluginMetadataFallback() {
  const source = await longFixture(2);
  const worker = installWorker({ peerMetadata: true, holdState: true, holdCompletion: true });
  const frames = holdFrames(true);
  const clock = holdPeerFallback({ allTimers: true });
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const deferred = spyOn(replicaHelpers, 'deferWorkerOpenReplica');
  let opened!: NonNullable<Awaited<ReturnType<OpenInWorker>>>;
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, source, layoutCompleteSession: null,
      onWorkerOpen: (worker) => { if (worker) opened = worker; },
    },
  });
  for (let turn = 0; turn < 1_000 && (!opened || result.current.host === null); turn += 1) {
    await act(async () => {
      clock.advance(0);
      await new Promise<void>((resolve) => setImmediate(resolve));
    });
  }
  await flushPluginFallback(clock);
  expect(result.current.host).not.toBeNull();
  expect(opened).toBeDefined();
  const session = result.current.core.session!;
  const metadata = observeMetadataBootstrap(session);
  expect(session.storyIds()).toEqual([]);
  expect(result.current.core.replicaReady).toBe(false);
  const replicaIndex = deferred.mock.calls.findIndex(([owner]) => owner === session);
  expect(replicaIndex).toBeGreaterThanOrEqual(0);
  const replica = deferred.mock.results[replicaIndex]!.value as ReturnType<typeof replicaHelpers.deferWorkerOpenReplica>;
  const snapshot = observePluginFallback(opened.documentRead({
    kind: 'readParagraphs', request: { view: 'accepted' },
  }), clock);
  await flushPluginFallback(clock);
  expect(snapshot.settled).toBe(true);
  expect(snapshot.error).toBeUndefined();
  if (!snapshot.value?.value.ok) throw new Error('The worker paragraph read did not succeed');
  const version = snapshot.value.version;
  const paragraphs = snapshot.value.value.paragraphs;
  expect(paragraphs.map((paragraph) => paragraph.text)).toEqual(['First paragraph', 'Tail paragraph']);
  const editor = result.current.pagedEditorRef.current!;
  const focus = mock(() => {});
  Object.assign(editor, { focus });
  const flush = spyOn(editor, 'flushPendingInput').mockImplementation(async () => {
    await replicaHelpers.awaitWorkerOpenReplica(session);
  });
  const start = spyOn(replica, 'start');
  const ensure = spyOn(replica, 'ensure');
  const requestReady = spyOn(replica, 'requestReady');
  const requested = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
  const ensured = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
  const readinessRequested = spyOn(replicaHelpers, 'requestWorkerOpenReplicaReadiness');
  const ready = spyOn(replicaHelpers, 'awaitWorkerOpenReplica');
  const load = spyOn(session, 'loadState');
  const apply = spyOn<YrsSession, 'applyEdits'>(session, 'applyEdits');
  const selection = spyOn(session, 'setSelection');
  const binding = testBinding();
  binding.state.admission = async () => {
    const flushed = await flushEditorInput(result.current.pagedEditorRef);
    if (!flushed.ok) throw new DocxCommandAdmissionError(flushed.code);
  };
  const commands = createDocxCommandController();
  commands.attach(binding.binding);
  const controller = new AbortController();
  const lifetime = new AbortController();
  const invocation: PluginInvocation<DocxPluginSnapshot> = {
    pluginId: 'acme.review',
    activation: {},
    snapshot: {} as DocxPluginSnapshot,
    signal: controller.signal,
    lifetimeSignal: lifetime.signal,
    state: () => null,
    setState: () => false,
    onCleanup: () => {},
    run: async () => {},
    commit: (write) => write(),
    refusal: () => controller.signal.aborted ? 'aborted' : null,
  };
  const queries = {
    sourceState: () => ({ status: 'ready' }),
    anchorRect: () => ({ pageIndex: 0, x: 0, y: 0, width: 1, height: 1 }),
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, version);
  const clients = createPluginClients(invocation, {
    pagedEditorRef: result.current.pagedEditorRef,
    writeMode: () => 'editing',
    viewer: () => false,
    commands: () => commands,
    layout: () => ({ queries, complete: false, failed: false }),
    subscribeLayout: (listener) => session.onUpdate(listener),
  }, () => ({ document: 'write', editBatches: true }), commands.store);
  return {
    worker, frames, clock, tasks, visibility, result, unmount, session, replica, metadata,
    version, paragraphs, clients, focus, flush, start, ensure, requestReady,
    requested, ensured, readinessRequested, ready, load, apply, selection, binding,
  };
}

async function expectPluginMetadataOwnerFallback(
  calls: (env: Awaited<ReturnType<typeof openingPluginMetadataFallback>>) => readonly Promise<unknown>[],
  complete: (values: readonly unknown[], env: Awaited<ReturnType<typeof openingPluginMetadataFallback>>) => void | Promise<void>
) {
  const env = await openingPluginMetadataFallback();
  const { clock, replica, result, session, worker } = env;
  let outcomes: ReturnType<typeof observePluginFallback>[] = [];
  try {
    expect(clock.now).toBe(0);
    expect(worker.posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(result.current.pagedEditorRef.current?.hasPendingInput()).toBe(false);
    outcomes = calls(env).map((call) => observePluginFallback(call, clock));
    await flushPluginFallback(clock);
    expectUnstarted();
    await act(async () => { clock.advance(9_999); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(9_999);
    expectUnstarted();
    await act(async () => { clock.advance(2); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(10_001);
    expect(replica.started).toBe(true);
    expect(env.start).toHaveBeenCalledTimes(1);
    expect(result.current.renderer.layoutCompleteSession).toBeNull();
    expect(env.requested.mock.calls).toEqual([[session]]);
    expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(env.load).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([]);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    expectPassiveCall();
    await act(async () => { worker.workers[0]!.release(); });
    await flushPluginFallback(clock);
    expect(env.tasks.tasks).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([]);
    await act(async () => { await env.tasks.run(); });
    await flushPluginFallback(clock);
    expect(env.load).not.toHaveBeenCalled();
    env.metadata.check();
    expect(env.tasks.tasks).toHaveLength(1);
    await act(async () => { await env.tasks.run(); });
    await flushPluginFallback(clock);
    expect(replica.hydrated).toBe(true);
    expect(replica.pending).toBe(true);
    const hydratedVersion = replica.readyVersion;
    expect(hydratedVersion).toBeDefined();
    if (hydratedVersion === undefined || replica.loadedVersion === undefined) {
      throw new Error('The replica did not record its hydrated versions');
    }
    expect(session.version()).toBe(hydratedVersion);
    expect(hydratedVersion).toBe(replica.loadedVersion);
    expect(hydratedVersion).not.toBe(env.version);
    expect(clock.now).toBe(10_001);
    expect(result.current.core.replicaReady).toBe(false);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    await act(async () => { clock.advance(2_999); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(13_000);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    await act(async () => { clock.advance(27_001); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(10_000 + 30_000 + 1);
    for (const outcome of outcomes) {
      expect(outcome.settled).toBe(true);
      expect(outcome.error).toBeUndefined();
      expect(outcome.at).toBeLessThanOrEqual(10_000 + 30_000 + 1);
    }
    expect(replica.pending).toBe(false);
    expect(result.current.core.replicaReady).toBe(true);
    const applied = env.apply.mock.results
      .flatMap((call) => call.type === 'return' && call.value ? [call.value] : [])
      .find((value) => value.ok && value.applied);
    expect(session.version()).toBe(applied?.version ?? hydratedVersion);
    expect(env.start).toHaveBeenCalledTimes(1);
    expect(env.requested.mock.calls).toEqual([[session]]);
    expect(env.load).not.toHaveBeenCalled();
    env.metadata.check();
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.core.session).toBe(session);
    expect<unknown[]>([null, session]).toContain(result.current.renderer.layoutCompleteSession);
    expect(result.current.errors).toEqual([]);
    expectPassiveCall();
    await complete(outcomes.map((outcome) => outcome.value), env);
  } finally {
    env.unmount();
    env.visibility.restore();
    env.tasks.restore();
    env.clock.restore();
    env.frames.restore();
  }

  function expectPassiveCall() {
    expect(env.ensure).not.toHaveBeenCalled();
    expect(env.requestReady).not.toHaveBeenCalled();
    expect(env.ensured).not.toHaveBeenCalled();
    expect(env.readinessRequested).not.toHaveBeenCalled();
    expect(env.ready.mock.calls).toEqual(outcomes.map(() => [session]));
    expect(env.flush).toHaveBeenCalledTimes(outcomes.length);
    expect(result.current.pagedEditorRef.current?.hasPendingInput()).toBe(false);
  }

  function expectUnstarted() {
    expect(replica.pending).toBe(true);
    expect(replica.started).toBe(false);
    expect(env.start).not.toHaveBeenCalled();
    expect(env.requested).not.toHaveBeenCalled();
    expect(env.load).not.toHaveBeenCalled();
    expect(env.apply).not.toHaveBeenCalled();
    expect(env.selection).not.toHaveBeenCalled();
    expect(env.focus).not.toHaveBeenCalled();
    expect(env.binding.calls).toEqual([]);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.renderer.layoutCompleteSession).toBeNull();
    expect(worker.posted.filter((request) => request.type === 'encodeState')).toEqual([]);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    expectPassiveCall();
  }
}

test('a pre-hydration worker version applies after unchanged owner hydration and becomes stale after an edit with metadata bootstrap', async () => {
  await expectPluginMetadataOwnerFallback(() => [], async (_values, env) => {
    const readyVersion = env.replica.readyVersion;
    if (readyVersion === undefined) throw new Error('The replica did not record its ready version');
    const request = {
      expectVersion: env.version,
      steps: [{
        op: 'replaceText' as const,
        target: { kind: 'paragraph' as const, story: 'body', paraId: env.paragraphs[0]!.paraId },
        text: 'Owner-ready edit',
      }],
    };
    expect(env.replica.handoverVersion).toBe(env.version);
    expect(env.session.version()).toBe(readyVersion);
    const applied = await env.clients.edits!.applyEdits(request);
    expect(applied).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
    expect(env.apply.mock.calls[0]![0].expectVersion).toBe(readyVersion);
    expect(env.session.paragraphs('body')[0]!.text).toBe('Owner-ready edit');
    const editedVersion = env.session.version();
    expect(editedVersion).not.toBe(env.replica.readyVersion);
    expect(await env.clients.edits!.applyEdits({
      ...request, steps: [{ ...request.steps[0]!, text: 'Stale edit' }],
    })).toMatchObject({ ok: false, version: editedVersion, failure: { code: 'stale-version' } });
    expect(env.apply.mock.calls[1]![0].expectVersion).toBe(env.version);
    expect(env.session.version()).toBe(editedVersion);
    expect(env.session.paragraphs('body')[0]!.text).toBe('Owner-ready edit');
    expect(await env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: env.paragraphs[1]!.paraId },
      { expectVersion: env.version, focus: true }
    )).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(env.selection).not.toHaveBeenCalled();
    expect(env.focus).not.toHaveBeenCalled();
    expect(env.result.current.searchReveals).toEqual([]);
  });
});

function parsedDocument(originalBuffer?: ArrayBuffer): Document {
  return {
    ...(originalBuffer ? { originalBuffer } : {}),
    package: {
      document: {
        content: [{
          type: 'paragraph',
          content: [{ type: 'run', content: [{ type: 'text', text: 'Host changed text' }] }],
        }],
      },
    },
  } as Document;
}

function useParsedDocumentHarness(props: { document: Document | null; viewer?: boolean }) {
  const history = useHistory<Document | null>(null);
  const [fontScope] = useState(() => {
    const scope = createFontLoadScope();
    scope.loadDocumentFonts = async () => {};
    scope.loadFontsWithMapping = async () => {};
    return scope;
  });
  useEffect(() => () => fontScope.dispose(), [fontScope]);
  const loadErrors = useRef<Error[]>([]);
  const commentsLoadedRef = useRef(false);
  const [loading, setLoadingState] = useState({ isLoading: false, parseError: null as string | null });
  const loader = useDocumentLoader({
    documentBuffer: null,
    initialDocument: props.document,
    workerViewer: props.viewer !== false,
    externalContent: false,
    history,
    pagedEditorRef: { current: null },
    setLoadingState,
    setComments: () => {},
    setShowCommentsSidebar: () => {},
    onError: (error) => loadErrors.current.push(error),
    resetForNewDocument: () => { commentsLoadedRef.current = false; },
    commentsLoadedRef,
    commentIdAllocator: createCommentIdAllocator(),
    setDocumentFonts: () => {},
    fontScope,
  });
  const harness = useHarness({
    ...initialProps, loader, viewer: props.viewer !== false, readOnly: props.viewer !== false,
  });
  return { ...harness, loader, history, loading, loadErrors: loadErrors.current };
}

for (const entry of ['prop', 'loadDocument'] as const) {
  test.each([false, true])(`a viewer opened by ${entry} serializes once with originalBuffer=%s and opens only in the worker`, async (originalBuffer) => {
    const { posted } = installWorker();
    const serialized = await longFixture(2);
    const parsed = parsedDocument(originalBuffer ? bytes.slice().buffer : undefined);
    const create = spyOn(docx, 'createDocx').mockResolvedValue(serialized.buffer as ArrayBuffer);
    const repack = spyOn(docx, 'repackDocx').mockResolvedValue(serialized.buffer as ArrayBuffer);
    const seeded = spyOn(yrsFacade, 'documentToYrs');
    const main = trackMainLoads();
    const { result, rerender, unmount } = renderHook(useParsedDocumentHarness, {
      initialProps: { document: entry === 'prop' ? parsed : null },
    });
    try {
      if (entry === 'loadDocument') act(() => result.current.ref.current!.loadDocument(parsed));
      await waitFor(() => expect(result.current.host).not.toBeNull());
      await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
      act(() => rerender({ document: entry === 'prop' ? parsed : null }));
      const opens = posted.filter((request) => request.type === 'open');
      expect(opens).toHaveLength(1);
      expect(new Uint8Array(opens[0].bytes)).toEqual(new Uint8Array(serialized));
      const writer = originalBuffer ? repack : create;
      expect(writer.mock.calls).toEqual([[parsed]]);
      expect(originalBuffer ? create : repack).not.toHaveBeenCalled();
      expect(seeded).not.toHaveBeenCalled();
      expect(workerOpenDocumentHeld(result.current.core.session!)).toBe(true);
      expect(result.current.core.session!.storyIds()).toEqual([]);
      expect(result.current.loader.yrsSeedDocument).toBeNull();
      expect(result.current.history.state).toBe(result.current.host!.document);
      expect(result.current.history.state).not.toBe(parsed);
      expect(result.current.mainOpens).toEqual([]);
      expect(result.current.errors).toEqual([]);
      expect(result.current.loadErrors).toEqual([]);
      expect(main.loads.length).toBeGreaterThan(0);
      for (const load of main.loads) expect(load.mock.calls).toHaveLength(0);
    } finally {
      unmount();
      main.restore();
      seeded.mockRestore();
      create.mockRestore();
      repack.mockRestore();
    }
  });

  test(`switching parsed viewer documents through ${entry} opens each serialized document in a new worker session`, async () => {
    const { posted } = installWorker();
    const replacement = await longFixture(2);
    const first = parsedDocument();
    const second = parsedDocument();
    const writer = spyOn(docx, 'createDocx')
      .mockResolvedValueOnce(bytes.slice().buffer as ArrayBuffer)
      .mockResolvedValueOnce(replacement.buffer as ArrayBuffer);
    const main = trackMainLoads();
    const { result, rerender, unmount } = renderHook(useParsedDocumentHarness, {
      initialProps: { document: entry === 'prop' ? first : null },
    });
    try {
      if (entry === 'loadDocument') act(() => result.current.ref.current!.loadDocument(first));
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const session = result.current.core.session;
      const generation = result.current.loader.yrsSeedGeneration;
      act(() => {
        if (entry === 'prop') rerender({ document: second });
        else result.current.ref.current!.loadDocument(second);
      });
      await waitFor(() => expect(result.current.core.sessionGeneration).toBe(generation + 1));
      const opens = posted.filter((request) => request.type === 'open');
      expect(opens).toHaveLength(2);
      expect(opens.map((request): Uint8Array => new Uint8Array(request.bytes))).toEqual([bytes, replacement]);
      expect(writer.mock.calls).toEqual([[first], [second]]);
      expect(result.current.core.session).not.toBe(session);
      expect(workerOpenDocumentHeld(result.current.core.session!)).toBe(true);
      expect(result.current.loader.yrsSeedDocument).toBeNull();
      expect(result.current.mainOpens).toEqual([]);
      expect(result.current.loadErrors).toEqual([]);
      for (const load of main.loads) expect(load.mock.calls).toHaveLength(0);
    } finally {
      unmount();
      main.restore();
      writer.mockRestore();
    }
  });

  test.each([false, true])(`a serialization failure through ${entry} reports a typed open error with originalBuffer=%s`, async (originalBuffer) => {
    const { posted } = installWorker();
    const cause = new Error('Cannot serialize the host document');
    const parsed = parsedDocument(originalBuffer ? bytes.slice().buffer : undefined);
    const writer = originalBuffer
      ? spyOn(docx, 'repackDocx').mockRejectedValue(cause)
      : spyOn(docx, 'createDocx').mockRejectedValue(cause);
    const main = trackMainLoads();
    const { result, unmount } = renderHook(useParsedDocumentHarness, {
      initialProps: { document: entry === 'prop' ? parsed : null },
    });
    try {
      if (entry === 'loadDocument') act(() => result.current.ref.current!.loadDocument(parsed));
      await waitFor(() => expect(result.current.loadErrors).toHaveLength(1));
      const error = result.current.loadErrors[0] as DocxWorkerError;
      expect(error).toBeInstanceOf(DocxWorkerError);
      expect(error.stage).toBe('open');
      expect(error.cause).toBe(cause);
      expect(writer.mock.calls).toEqual([[parsed]]);
      expect(result.current.loading).toEqual({ isLoading: false, parseError: error.message });
      expect(result.current.core.session).toBeNull();
      expect(result.current.history.state).toBeNull();
      expect(result.current.loader.yrsSeedDocument).toBeNull();
      expect(result.current.loader.yrsSeedBytes).toBeNull();
      expect(posted.filter((request) => request.type === 'open')).toEqual([]);
      expect(result.current.mainOpens).toEqual([]);
      for (const load of main.loads) expect(load.mock.calls).toHaveLength(0);
    } finally {
      unmount();
      main.restore();
      writer.mockRestore();
    }
  });
}

test.each(['prop', 'loadDocument'] as const)('an editor opened by %s keeps its parsed document and eager main session without serialization', async (entry) => {
  const { posted } = installWorker();
  const parsed = parsedDocument(bytes.slice().buffer);
  const create = spyOn(docx, 'createDocx');
  const repack = spyOn(docx, 'repackDocx');
  const { result, unmount } = renderHook(useParsedDocumentHarness, {
    initialProps: { document: entry === 'prop' ? parsed : null, viewer: false },
  });
  try {
    if (entry === 'loadDocument') act(() => result.current.ref.current!.loadDocument(parsed));
    await waitFor(() => expect(result.current.core.session).not.toBeNull());
    expect(result.current.core.session!.paragraphs('body').some((paragraph) => paragraph.text === 'Host changed text')).toBe(true);
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.loader.yrsSeedDocument).toBe(parsed);
    expect(result.current.history.state).toBe(parsed);
    expect(posted.filter((request) => request.type === 'open')).toEqual([]);
    expect(create).not.toHaveBeenCalled();
    expect(repack).not.toHaveBeenCalled();
    expect(result.current.loadErrors).toEqual([]);
  } finally {
    unmount();
    create.mockRestore();
    repack.mockRestore();
  }
});

test.each(['success', 'failure'] as const)('a replaced parsed viewer serialization ignores a late %s', async (outcome) => {
  const { posted } = installWorker();
  let resolve!: (buffer: ArrayBuffer) => void;
  let reject!: (error: Error) => void;
  const pending = new Promise<ArrayBuffer>((yes, no) => { resolve = yes; reject = no; });
  const first = parsedDocument();
  const second = parsedDocument();
  const writer = spyOn(docx, 'createDocx')
    .mockImplementationOnce(() => pending)
    .mockResolvedValueOnce(bytes.slice().buffer as ArrayBuffer);
  const main = trackMainLoads();
  const { result, rerender, unmount } = renderHook(useParsedDocumentHarness, {
    initialProps: { document: first },
  });
  try {
    expect(writer).toHaveBeenCalledTimes(1);
    act(() => rerender({ document: second }));
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session;
    await act(async () => {
      if (outcome === 'success') resolve(bytes.slice().buffer as ArrayBuffer);
      else reject(new Error('The replaced document could not serialize'));
      await pending.catch(() => {});
    });
    expect(writer.mock.calls).toEqual([[first], [second]]);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(result.current.core.session).toBe(session);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.loadErrors).toEqual([]);
    for (const load of main.loads) expect(load.mock.calls).toHaveLength(0);
  } finally {
    unmount();
    main.restore();
    writer.mockRestore();
  }
});

test('a hydrated editor exports ordinary and pinned pages through the real layout request path', async () => {
  const { posted } = installWorker();
  const { result } = renderHook(useHarness, { initialProps: { ...initialProps } });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  act(() => result.current.pipeline.runLayoutPipeline());
  await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
  act(() => result.current.presentFrame());
  await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
  await act(async () => { await result.current.renderer.settledDisplayList(null, 3000, 'window'); });
  expect(await result.current.pipeline.readLayoutRequest()).toBeNull();
  expect(result.current.pipeline.getLayoutRequest()).not.toBeNull();
  const session = result.current.core.session!;
  const peerExport = spyOn(session, 'exportStructuredWithPagesFor');
  try {
    let first!: Awaited<ReturnType<DocxEditorRef['exportStructuredWithPages']>>;
    await act(async () => { first = await result.current.ref.current!.exportStructuredWithPages({ revisionView: 'markup' }); });
    if (!first.ok) throw new Error(first.failure.message);
    expect(first.version).toBe(session.version());
    const options = { revisionView: 'markup' as const, expectLayoutVersion: first.content.layout.layoutVersion };
    const beforePinned = posted.length;
    let pinned!: Awaited<ReturnType<DocxEditorRef['exportStructuredWithPages']>>;
    await act(async () => {
      pinned = await result.current.ref.current!.exportStructuredWithPages(options);
    });
    expect(pinned).toEqual(first);
    expect(posted.slice(beforePinned).filter((request) => request.type === 'documentRead' && request.read.kind === 'exportStructuredWithPages')).toHaveLength(1);
    expect(posted.slice(beforePinned).some(({ type }) => type === 'bootstrap' || type === 'sync' || type === 'completeLayout')).toBe(false);
    expect(peerExport).not.toHaveBeenCalled();
  } finally {
    peerExport.mockRestore();
  }
});

test('main-thread takeover retires worker exports while an in-flight export still rejects', async () => {
  const source = await longFixture(1);
  let holdReads = false;
  const { workers, posted, replies } = installWorker({ holdReply: (request) => holdReads && request.type === 'documentRead' });
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, source },
  });
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  let peerExport: ReturnType<typeof spyOn<YrsSession, 'exportStructuredWithPagesFor'>> | undefined;
  try {
    await frames.waitFor(() => expect(result.current.host).not.toBeNull());
    act(() => result.current.pipeline.runLayoutPipeline());
    await frames.waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    act(() => result.current.presentFrame());
    await frames.waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, 3000, 'window'), 3000);
    const session = result.current.core.session!;
    const api = result.current.ref.current!;
    const first = await frames.untilCommitted(api.exportStructuredWithPages({ revisionView: 'markup' }), 3000);
    if (!first.ok) throw new Error(first.failure.message);
    expect(workerOpenExport(session)).not.toBeNull();
    const request = result.current.pipeline.getLayoutRequest()!;
    session.layoutDocumentWithRegionsRetainedJson(request);
    session.buildDisplayListFrame('{}', 0);
    session.setSelection({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 });
    peerExport = spyOn(session, 'exportStructuredWithPagesFor');
    const previousId = posted.at(-1)!.id;
    holdReads = true;
    const exporting = api.exportStructuredWithPages({ revisionView: 'markup', expectLayoutVersion: first.content.layout.layoutVersion });
    void exporting.catch(() => {});
    await frames.waitFor(() => {
      const request = posted.find((request) => request.type === 'documentRead' && request.id > previousId);
      expect(request).toBeDefined();
      expect(replies.has(request!.id)).toBe(true);
    });
    workers[0]!.hold();
    let input!: ReturnType<typeof result.current.renderer.applyInput>;
    act(() => { input = result.current.renderer.applyInput('Recovered '); });
    await frames.waitFor(() => expect(posted.some((request) => request.type === 'applyInput')).toBe(true));
    act(() => { workers[0]!.onerror?.({ message: 'worker crashed' } as ErrorEvent); });
    await frames.untilCommitted(Promise.all([input, expect(exporting).rejects.toThrow(ResidentWorkerFailureError)]), 3000);
    expect(workerOpenExport(session)).toBeNull();
    expect(peerExport).not.toHaveBeenCalled();
    const exported = await frames.untilCommitted(api.exportStructuredWithPages({ revisionView: 'markup' }), 3000);
    if (!exported.ok) throw new Error(exported.failure.message);
    expect(exported.version).toBe(session.version());
    expect(JSON.stringify(exported.content.structured)).toContain('Recovered ');
    expect(peerExport).toHaveBeenCalled();
    expect(workers).toHaveLength(1);
  } finally {
    unmount();
    peerExport?.mockRestore();
    errorLog.mockRestore();
    frames.restore();
    globalThis.Worker = originalWorker;
  }
});

test('registered editor export reconciles worker repairs without echoing updates to the worker', async () => {
  const { workers, posted } = installWorker();
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, { initialProps: { ...initialProps } });
  let applyLocal: ReturnType<typeof spyOn<YrsSession, 'applyLocalUpdate'>> | undefined;
  let peerExport: ReturnType<typeof spyOn<YrsSession, 'exportStructuredWithPagesFor'>> | undefined;
  try {
    await frames.waitFor(() => expect(result.current.host).not.toBeNull());
    act(() => result.current.pipeline.runLayoutPipeline());
    await frames.waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    act(() => result.current.presentFrame());
    await frames.waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, 3000, 'window'), 3000);
    const session = result.current.core.session!;
    applyLocal = spyOn(session, 'applyLocalUpdate');
    peerExport = spyOn(session, 'exportStructuredWithPagesFor');
    workers[0]!.sessions[0]!.applyRawOps('body', [{ op: 'insert', index: 0, text: 'Worker repair ' }]);
    const before = posted.filter(({ type }) => type === 'applyUpdate').length;
    const beforeExport = posted.length;
    const exported = await frames.untilCommitted(result.current.ref.current!.exportStructuredWithPages({ revisionView: 'markup' }), 3000);
    if (!exported.ok) throw new Error(exported.failure.message);
    expect(exported.version).toBe(session.version());
    expect(exported.content.layout.documentVersion).toBe(session.version());
    expect(JSON.stringify(exported.content.structured)).toContain('Worker repair ');
    expect(session.paragraphs('body')[0]!.text).toStartWith('Worker repair ');
    expect(applyLocal).toHaveBeenCalled();
    expect(posted.filter(({ type }) => type === 'applyUpdate')).toHaveLength(before);
    expect(posted.slice(beforeExport).some(({ type }) => type === 'syncUpdate')).toBe(true);
    expect(posted.slice(beforeExport).some(({ type }) => type === 'sync')).toBe(true);
    expect(peerExport).not.toHaveBeenCalled();
  } finally {
    unmount();
    applyLocal?.mockRestore();
    peerExport?.mockRestore();
    frames.restore();
    globalThis.Worker = originalWorker;
  }
});

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
    harness = useHarness({ ...initialProps, source });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        onLayoutComputed={harness.renderer.onLayoutComputed}
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

test('saved paragraph ID claims refresh editor point geometry without another edit', async () => {
  installWorker();
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const source = await longFixture(1);
  const editor = createRef<PagedEditorRef>();
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function Editable() {
    harness = useHarness({ ...initialProps, source, viewer: false });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        onLayoutComputed={(layout, session) => harness.renderer.onLayoutComputed(layout, session)}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries} />
    </>;
  }
  render(<Editable />);
  await waitFor(() => expect(harness.renderer.status).toBe('ready'));
  act(() => harness.presentFrame());
  await waitFor(() => expect(harness.core.replicaReady).toBe(true));
  const session = harness.core.session!;
  act(() => {
    session.splitParagraph({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 5 });
    editor.current!.syncYrsInputState(true, ['body']);
  });
  await waitFor(() => expect(sourceVersionOf(harness.renderer.queries)).toBe(session.version()));
  await act(async () => { await harness.renderer.settledDisplayList(null, 3000); });
  const point = () => {
    const queries = harness.renderer.queries!;
    const size = queries.pageSize(0)!;
    canvasHost.current!.firstElementChild!.getBoundingClientRect = () => ({
      left: 0, top: 0, right: size.width, bottom: size.height, ...size,
    }) as DOMRect;
    markPresented(canvasHost.current!, queries.displayList);
    const caret = queries.caretRect(1)!;
    return editor.current!.getPositionAtPoint(caret.x, caret.y + caret.height / 2);
  };
  expect(point()).not.toBeNull();
  const beforeSave = session.encodeStateVector();
  await act(async () => {
    expect(await workerOpenSave(session)!.save([], session)).toBeInstanceOf(ArrayBuffer);
  });
  expect(session.encodeStateVector()).not.toEqual(beforeSave);
  await waitFor(() => expect(sourceVersionOf(harness.renderer.queries)).toBe(session.version()));
  await act(async () => { await harness.renderer.settledDisplayList(null, 3000); });
  expect(point()).not.toBeNull();
  expect(harness.errors).toEqual([]);
}, 20_000);

test("a worker save's own paragraph ID claims leave no story dirty for the next save", async () => {
  const { posted } = installWorker();
  const savedStories = () => posted.filter((request) => request.type === 'save').map((request) => request.stories);
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const source = await longFixture(1);
  const editor = createRef<PagedEditorRef>();
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function Editable() {
    harness = useHarness({ ...initialProps, source, viewer: false });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        onLayoutComputed={(layout, session) => harness.renderer.onLayoutComputed(layout, session)}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries} />
    </>;
  }
  render(<Editable />);
  await waitFor(() => expect(harness.renderer.status).toBe('ready'));
  act(() => harness.presentFrame());
  await waitFor(() => expect(harness.core.replicaReady).toBe(true));
  const session = harness.core.session!;
  act(() => {
    session.splitParagraph({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 5 });
    editor.current!.syncYrsInputState(true, ['body']);
  });
  await waitFor(() => expect(sourceVersionOf(harness.renderer.queries)).toBe(session.version()));
  await act(async () => { await harness.renderer.settledDisplayList(null, 3000); });
  const beforeSave = session.encodeStateVector();
  await act(async () => {
    expect(await workerOpenSave(session)!.save([], session)).toBeInstanceOf(ArrayBuffer);
  });
  expect(session.encodeStateVector()).not.toEqual(beforeSave);
  expect(savedStories()).toEqual([['body']]);
  await act(async () => {
    expect(await workerOpenSave(session)!.save([], session)).toBeInstanceOf(ArrayBuffer);
  });
  expect(savedStories()).toEqual([['body'], []]);
  act(() => {
    session.splitParagraph({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 2 });
    editor.current!.syncYrsInputState(true, ['body']);
  });
  await waitFor(() => expect(sourceVersionOf(harness.renderer.queries)).toBe(session.version()));
  await act(async () => { await harness.renderer.settledDisplayList(null, 3000); });
  await act(async () => {
    expect(await workerOpenSave(session)!.save([], session)).toBeInstanceOf(ArrayBuffer);
  });
  expect(savedStories()).toEqual([['body'], [], ['body']]);
  expect(harness.errors).toEqual([]);
}, 20_000);

test('overlapping editor worker saves capture their stories in call order', async () => {
  const { workers, posted, received, reply } = installWorker({ holdReply: (request) => request.type === 'save' });
  const savedStories = () => posted.filter((request) => request.type === 'save').map((request) => request.stories);
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const source = await longFixture(1);
  const editor = createRef<PagedEditorRef>();
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function Editable() {
    harness = useHarness({ ...initialProps, source, viewer: false });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        onLayoutComputed={(layout, session) => harness.renderer.onLayoutComputed(layout, session)}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries} />
    </>;
  }
  render(<Editable />);
  await waitFor(() => expect(harness.renderer.status).toBe('ready'));
  act(() => harness.presentFrame());
  await waitFor(() => expect(harness.core.replicaReady).toBe(true));
  const session = harness.core.session!;
  act(() => {
    session.splitParagraph({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 5 });
    editor.current!.syncYrsInputState(true, ['body']);
  });
  await waitFor(() => expect(sourceVersionOf(harness.renderer.queries)).toBe(session.version()));
  await act(async () => { await harness.renderer.settledDisplayList(null, 3000); });
  const events: string[] = [];
  const worker = workers[0]!;
  const postMessage = worker.postMessage.bind(worker);
  const send = spyOn(worker, 'postMessage').mockImplementation((request, transfer) => {
    if (request.type === 'save') events.push('request');
    postMessage(request, transfer);
  });
  try {
    await act(async () => {
      const save = workerOpenSave(session)!;
      const first = save.save([], session).then((bytes) => {
        events.push('resolved');
        return bytes;
      });
      const second = save.save([], session);
      const firstRequest = await received('save');
      expect(savedStories()).toEqual([['body']]);
      expect(events).toEqual(['request']);
      reply(firstRequest);
      const secondRequest = await received('save', firstRequest.id);
      expect(events).toEqual(['request', 'resolved', 'request']);
      expect(savedStories()).toEqual([['body'], []]);
      reply(secondRequest);
      for (const bytes of await Promise.all([first, second])) {
        expect(bytes).toBeInstanceOf(ArrayBuffer);
        expect(bytes.byteLength).toBeGreaterThan(0);
      }
    });
    expect(harness.errors).toEqual([]);
  } finally {
    send.mockRestore();
  }
}, 20_000);

test('a viewer command preserves select-all while its worker read is pending', async () => {
  let heldUnit = false;
  const { posted, replies, responses, reply } = installWorker({
    holdReply: (request) => {
      if (heldUnit || request.type !== 'documentRead' || request.read.kind !== 'selectionUnit') return false;
      heldUnit = true;
      return true;
    },
  });
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const source = await longFixture(2);
  const editor = createRef<PagedEditorRef>();
  const bridge = { current: null as PagedEditorCommandBridge | null };
  let harness!: ReturnType<typeof useHarness>;
  function ReadOnly() {
    harness = useHarness({ ...initialProps, source, readOnly: true, viewer: true, workerProposals: true });
    return <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core} readOnly
      viewerDocumentRead={harness.renderer.readWorkerDocument}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
      fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
      layoutInWorker={harness.renderer.layoutInWorker}
      displayListQueries={harness.renderer.queries}
      commandBridgeRef={bridge} />;
  }
  const view = render(<ReadOnly />);
  try {
    await waitFor(() => expect(harness.renderer.status).toBe('ready'));
    await waitFor(() => expect(harness.core.workerProposalsReady).toBe(true));
    act(() => harness.pipeline.runLayoutPipeline());
    await waitFor(() => expect(presentedWorkerVersion(harness.renderer.queries)).toBe(harness.core.session!.version()));
    act(() => harness.presentFrame());
    const session = harness.core.session!;
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    expect(textarea.readOnly).toBe(true);
    fireEvent.keyDown(textarea, { key: 'a', ctrlKey: true });
    let unit!: ResidentEngineWorkerRequest;
    await waitFor(() => {
      unit = posted.find((request) => request.type === 'documentRead' && request.read.kind === 'selectionUnit')!;
      expect(unit && replies.has(unit.id)).toBe(true);
    });
    expect(unit).toMatchObject({ read: { expectVersion: session.version() } });
    expect(responses.get(unit)).toMatchObject({ ok: true, read: { version: session.version(), value: { anchor: 0, head: 33 } } });
    const operation = mock(() => editor.current!.readSelectedText());
    let commandSettled = false;
    const command = bridge.current!.runAfterPendingInput(operation).then((text) => {
      commandSettled = true;
      return text;
    });
    await act(async () => {});
    expect(operation).not.toHaveBeenCalled();
    expect(bridge.current!.hasPendingInput()).toBe(true);
    expect(commandSettled).toBe(false);
    let selected!: string | null;
    await act(async () => { reply(unit); });
    expect(editor.current!.getSelectionRange()).toEqual({ from: 0, to: 33 });
    selected = await command;
    expect(selected).toBe('First paragraph\nTail paragraph');
    expect(operation).toHaveBeenCalledTimes(1);
    expect(bridge.current!.hasPendingInput()).toBe(false);
    expect(workerOpenDocumentHeld(session)).toBe(true);
    expect(harness.core.replicaReady).toBe(false);
    expect(harness.mainOpens).toEqual([]);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(harness.errors).toEqual([]);
  } finally {
    view.unmount();
  }
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
    harness = useHarness({ ...initialProps, source });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        onLayoutComputed={harness.renderer.onLayoutComputed}
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

function holdFrames(holdPeerIdle = false) {
  const request = globalThis.requestAnimationFrame;
  const cancel = globalThis.cancelAnimationFrame;
  const idle = holdPeerIdle ? holdIdle() : {
    callbacks: new Map<number, { callback: IdleRequestCallback; options?: IdleRequestOptions }>(),
    run: () => {},
    restore: () => {},
  };
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
    runIdle: idle.run,
    idleCallbacks: idle.callbacks,
    async waitFor(assertion: () => void, timeoutMs = 1000): Promise<void> {
      const deadline = Date.now() + timeoutMs;
      for (;;) {
        try {
          assertion();
          return;
        } catch (error) {
          if (Date.now() >= deadline) throw error;
        }
        await act(async () => {
          run();
          await new Promise<void>((resolve) => setTimeout(resolve, 0));
        });
      }
    },
    async until<T>(promise: Promise<T>): Promise<T> {
      let settled = false;
      void promise.then(() => { settled = true; }, () => { settled = true; });
      while (!settled) {
        run();
        await new Promise<void>((resolve) => setImmediate(resolve));
      }
      return promise;
    },
    async untilCommitted<T>(promise: Promise<T>, timeoutMs?: number): Promise<T> {
      let settled = false;
      void promise.then(() => { settled = true; }, () => { settled = true; });
      const deadline = timeoutMs === undefined ? Infinity : Date.now() + timeoutMs;
      while (!settled) {
        if (Date.now() >= deadline) throw new Error(`Operation did not settle within ${timeoutMs}ms`);
        await act(async () => {
          run();
          await new Promise<void>((resolve) => setImmediate(resolve));
        });
      }
      await act(async () => {});
      return promise;
    },
    async settleAndIdle(promise: Promise<unknown>) {
      await this.untilCommitted(promise);
      await act(async () => idle.run());
    },
    restore: registerRestore(() => {
      globalThis.requestAnimationFrame = request;
      globalThis.cancelAnimationFrame = cancel;
      idle.restore();
    }),
  };
}

function holdReplicaTimers() {
  const schedule = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  const pending = new Map<ReturnType<typeof setTimeout>, { at: number; run: () => void }>();
  let now = 0;
  let nextId = 0;
  globalThis.setTimeout = ((...input: Parameters<typeof setTimeout>) => {
    const [callback, delay, ...args] = input;
    if ((delay === 1000 || delay === 5000 || delay === 10_000) && typeof callback === 'function') {
      const id = --nextId as unknown as ReturnType<typeof setTimeout>;
      pending.set(id, { at: now + delay, run: () => callback(...args) });
      return id;
    }
    return schedule(callback, delay, ...args);
  }) as typeof setTimeout;
  globalThis.clearTimeout = ((id: Parameters<typeof clearTimeout>[0]) => {
    if (id !== undefined && pending.delete(id as ReturnType<typeof setTimeout>)) return;
    cancel(id);
  }) as typeof clearTimeout;
  return {
    pending,
    advance(milliseconds: number) {
      now += milliseconds;
      for (const [id, timer] of [...pending]) {
        if (timer.at > now) continue;
        pending.delete(id);
        timer.run();
      }
    },
    restore: registerRestore(() => {
      globalThis.clearTimeout = cancel;
      globalThis.setTimeout = schedule;
    }),
  };
}

function expectLoadingMutations(api: DocxEditorRef, editor: PagedEditorRef, paraId = '00000001') {
  const session = editor.getYrsSession()!;
  const started = replicaHelpers.workerOpenReplicaStarted(session);
  const request = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
  const requests = request.mock.calls.length;
  const ensure = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
  const admissions = ensure.mock.calls.length;
  const calls: Partial<Record<keyof DocxEditorRef, unknown[]>> = {
    addComment: [{ paraId: '00000001', search: 'paragraph', text: 'Check', author: 'Ann' }],
    proposeChange: [{ paraId: '00000001', search: 'paragraph', replaceWith: 'text', author: 'Host' }],
    applyFormatting: [{ paraId: '00000001', search: 'paragraph', marks: { bold: true } }],
    setParagraphStyle: [{ paraId: '00000001', styleId: 'Normal' }],
    insertBreak: [{ paraId: '00000001', type: 'page' }],
  };
  for (const member of DOCX_REF_REPLICA_LOADING_MUTATIONS) {
    const args = calls[member];
    if (!args) throw new Error(`Missing loading mutation: ${member}`);
    expect(() => Reflect.apply(api[member] as Function, api, args)).toThrow(DocxReplicaNotReadyError);
  }
  const navigate = spyOn(editor, 'scrollToParaId');
  expect(api.getDocument()).toBeNull();
  expect(api.getEditorRef()).toBeNull();
  expect(api.findInDocument('paragraph')).toEqual([]);
  expect(api.scrollToParaId(paraId)).toBe(false);
  expect(navigate).not.toHaveBeenCalled();
  expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
  expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(started);
  expect(request.mock.calls).toHaveLength(requests);
  expect(ensure.mock.calls).toHaveLength(admissions);
}

function expectReadyMutations(api: DocxEditorRef, session: YrsSession, editor: PagedEditorRef, reveals: readonly number[], paraId = session.paragraphs('body')[0]!.paraId) {
  expect(api.getDocument()).not.toBeNull();
  expect(api.getEditorRef()).toBe(editor);
  const first = session.paragraphs('body')[0]!;
  expect(api.findInDocument(first.text)).toContainEqual(expect.objectContaining({ paraId: first.paraId, match: first.text }));
  const revealed = reveals.length;
  expect(api.scrollToParaId(paraId)).toBe(true);
  const target = editor.yrsLocToDisplayPosition({ story: 'body', paraId, offset: 0 });
  expect(target).not.toBeNull();
  expect(reveals.slice(revealed)).toEqual([target!]);
  expect(reveals).toContain(target!);
  const last = session.paragraphs('body').at(-1)!;
  const search = 'Replica contract';
  let mutationParaId!: string;
  act(() => {
    const span = session.locateParagraph('body', last.paraId);
    mutationParaId = session.splitParagraph({ story: 'body', paraId: last.paraId, offset: span.end - span.start }).secondParaId;
    session.insertText({ story: 'body', paraId: mutationParaId, offset: 0 }, search);
    session.setParagraphAttr(mutationParaId, 'pStyle', 'Heading1');
  });
  act(() => {
    const comment = api.addComment({ paraId: mutationParaId, search, text: 'Check', author: 'Ann' });
    expect(comment).toEqual(expect.any(Number));
    expect(session.resolveComment(String(comment))).toContainEqual(expect.objectContaining({ story: 'body' }));
    expect(api.applyFormatting({ paraId: mutationParaId, search, marks: { bold: true } })).toBe(true);
    expect(session.storySegments('body')).toContainEqual(expect.objectContaining({
      kind: 'text', text: search, attributes: expect.objectContaining({ bold: true }),
    }));
    expect(api.setParagraphStyle({ paraId: mutationParaId, styleId: 'Normal' })).toBe(true);
    expect(session.paragraphs('body').find((entry) => entry.paraId === mutationParaId)!.properties.pStyle).toBe('Normal');
    expect(api.proposeChange({ paraId: mutationParaId, search: '', replaceWith: 'Replica ready', author: 'Host' })).toBe(true);
    expect(session.listRevisions()).toContainEqual(expect.objectContaining({
      kind: 'insertion', author: 'Host', preview: 'Replica ready',
    }));
    const breaks = session.storySegments('body').filter((entry) => entry.kind === 'embed' && entry.embedKind === 'pageBreak').length;
    expect(api.insertBreak({ paraId: mutationParaId, type: 'page' })).toBe(true);
    expect(session.storySegments('body').filter((entry) => entry.kind === 'embed' && entry.embedKind === 'pageBreak')).toHaveLength(breaks + 1);
  });
}

function trackMainLoads() {
  const loads: Array<{ mock: { calls: readonly unknown[] }; mockRestore(): void }> = [];
  const factory = spyOn(yrsFacade, 'createYrsSession').mockImplementation(async (options) => {
    const session = await originalCreateYrsSession(options);
    for (const method of ['openDocx', 'openDocxPreview', 'loadState', 'applyUpdate'] as const) {
      loads.push(spyOn(session, method));
    }
    return session;
  });
  return {
    loads,
    restore: registerRestore(() => {
      factory.mockRestore();
      for (const load of loads) load.mockRestore();
    }),
  };
}

test('textarea focus leaves an eager editor replica waiting for its frame', async () => {
  const { posted } = installWorker({ holdState: true });
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, { initialProps });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const core = result.current.core;
    const view = render(
      <YrsInput
        enabled
        readOnly={false}
        session={core.session}
        replicaReadyRef={core.replicaReadyRef}
        inputPositionMap={core.inputPositionMap}
        displayPositionToLoc={() => null}
        locToDisplayPosition={() => null}
        onStateChange={() => {}}
        onDirectInput={() => {}}
      />
    );
    const textarea = view.getByTestId('yrs-input');
    act(() => {
      textarea.focus();
      fireEvent.keyDown(textarea, { key: 'ArrowRight' });
    });
    await act(async () => {});
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
  } finally {
    unmount();
    cleanup();
    frames.restore();
  }
});

test.each([true, false])('worker hydration retains comment authors and dates with readOnly=%s', async (readOnly) => {
  installWorker();
  const source = new Uint8Array(readFileSync(resolve(
    import.meta.dir,
    '../../../../../../crates/docx-edit/tests/fixtures/structured-export/principal.docx'
  )));
  const seeded = await createYrsSession();
  sessions.push(seeded);
  seeded.openDocx(source, true);
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, source, readOnly },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    await act(async () => { await requestWorkerOpenReplica(session); });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
    const options = { revisionView: 'markup' as const, stories: ['comments' as const] };
    const expected = seeded.exportStructured(options);
    const actual = session.exportStructured(options);
    if (!expected.ok) throw new Error(expected.failure.message);
    if (!actual.ok) throw new Error(actual.failure.message);
    const metadata = (content: typeof expected.content) => content.stories.map((story) => ({
      author: story.comment?.author,
      date: story.comment?.date,
    }));
    expect(metadata(expected.content)).toContainEqual({ author: expect.any(String), date: expect.any(String) });
    expect(metadata(actual.content)).toEqual(metadata(expected.content));
  } finally {
    unmount();
  }
});

test('a viewer holds its document without a deferred replica through revisions, reads and idle', async () => {
  const { posted } = installWorker({ revisionCount: 1 });
  const frames = holdFrames();
  const timers = holdReplicaTimers();
  const main = trackMainLoads();
  const deferred = spyOn(replicaHelpers, 'deferWorkerOpenReplica');
  const requested = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
  const ensured = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
  const revisions = mock(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, viewer: true, readOnly: true, onWorkerRevisions: revisions },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    expect(workerOpenDocumentHeld(session)).toBe(true);
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    await waitFor(() => expect(revisions).toHaveBeenCalledTimes(1));
    act(() => {
      result.current.presentFrame();
      result.current.core.requestReplica();
      result.current.openCommentsSidebar();
      result.current.core.scheduleCompatibilityWarm();
      expect(result.current.core.failOpening(new DocxWorkerError('layout'), session)).toBe(false);
      expect(result.current.core.documentFromYrs()).toBeNull();
      frames.run();
      frames.run();
      timers.advance(6000);
    });
    await act(async () => {});
    expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.core.replicaReadyRef?.current).toBe(false);
    expect(session.storyIds()).toEqual([]);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([]);
    expect(deferred).not.toHaveBeenCalled();
    expect(requested).not.toHaveBeenCalled();
    expect(ensured).not.toHaveBeenCalled();
    expect(main.loads.length).toBeGreaterThan(0);
    for (const load of main.loads) expect(load.mock.calls).toHaveLength(0);
  } finally {
    unmount();
    deferred.mockRestore();
    requested.mockRestore();
    ensured.mockRestore();
    main.restore();
    timers.restore();
    frames.restore();
  }
});

test.each([false, true])('viewer revision discovery never loads a replica with failed count=%s', async (failed) => {
  const { posted } = installWorker({ revisionCount: 1, failRevisionCount: failed });
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, viewer: true, readOnly: true },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    await waitFor(() => expect(posted.some((request) => request.type === 'revisionCount')).toBe(true));
    if (failed) {
      await waitFor(() => expect(result.current.errors).toHaveLength(1));
      expect(result.current.errors[0]).toBeInstanceOf(DocxWorkerError);
    } else {
      await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    }
    await act(async () => {});
    expect(workerOpenDocumentHeld(result.current.core.session!)).toBe(true);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
  } finally {
    unmount();
  }
});

test.each([false, true])('leaving viewer kind loads the editor replica once with a frame=%s', async (painted) => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames();
  const timers = holdReplicaTimers();
  const props = { ...initialProps, viewer: true };
  const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    if (painted) {
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
      act(() => result.current.presentFrame());
    }
    expect(workerOpenDocumentHeld(session)).toBe(true);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => timers.advance(6000));
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => rerender({ ...props, viewer: false }));
    expect(result.current.core.session).toBe(session);
    expect(workerOpenDocumentHeld(session)).toBe(false);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    act(() => {
      if (painted) {
        frames.run();
        frames.run();
      } else {
        timers.advance(10_000);
      }
    });
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    await act(async () => { workers[0].release(); await awaitWorkerOpenReplica(session); });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    act(() => {
      frames.run();
      frames.run();
      timers.advance(6000);
    });
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    timers.restore();
    frames.restore();
  }
});

test.each(['refused', 'unavailable', 'failed'] as const)('a %s viewer preview skips the main preview and opens the full document in the worker', async (outcome) => {
  const { posted } = installWorker({ refusePreview: outcome === 'refused' });
  const main = trackMainLoads();
  const deferred = spyOn(replicaHelpers, 'deferWorkerOpenReplica');
  const openPreviewInWorker = outcome === 'refused' ? undefined : mock(async () => {
    if (outcome === 'failed') throw new Error('worker preview failed');
    return null;
  });
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, source: longBytes, viewer: true, readOnly: true,
      previewFirstPage: true, workerPreview: true, openPreviewInWorker,
    },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    expect(result.current.core.previewing).toBe(false);
    expect(workerOpenDocumentHeld(result.current.core.session!)).toBe(true);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([]);
    expect(deferred).not.toHaveBeenCalled();
    expect(main.loads.length).toBeGreaterThan(0);
    for (const load of main.loads) expect(load.mock.calls).toHaveLength(0);
    expect(posted.filter((request) => request.type === 'open' && request.previewBlocks === undefined)).toHaveLength(1);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  } finally {
    unmount();
    deferred.mockRestore();
    main.restore();
  }
});

test.each([false, true])('viewer save keeps the document held and worker rendering with proposals=%s', async (workerProposals) => {
  const { posted } = installWorker();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, viewer: true, readOnly: true, workerProposals },
  });
  await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
  const session = result.current.core.session!;
  const saved: ArrayBuffer[] = [];
  const errors: Error[] = [];
  const release = spyOn(replicaHelpers, 'releaseWorkerOpenDocument');
  const io = renderHook(() => useFileIO({
    pagedEditorRef: result.current.pagedEditorRef,
    viewerSession: true,
    resolveImage: () => null,
    comments: [],
    documentName: undefined,
    onSave: (buffer) => saved.push(buffer),
    downloadOnSave: false,
    onError: (error) => errors.push(error),
    onOpen: undefined,
    onPrint: undefined,
    onDocumentNameChange: undefined,
    loadBuffer: async () => {},
    focusActiveEditor: () => {},
  }));
  const project = spyOn(result.current.pagedEditorRef.current!, 'getDocument');
  try {
    expect(workerOpenDocumentHeld(session)).toBe(true);
    expect(result.current.core.documentFromYrs()).toBeNull();
    await act(async () => { await io.result.current.handleSave(); });
    expect(saved).toHaveLength(1);
    expect(errors).toEqual([]);
    expect(release).not.toHaveBeenCalled();
    expect(project).not.toHaveBeenCalled();
    expect(workerOpenDocumentHeld(session)).toBe(true);
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(posted.filter((request) => request.type === 'save')).toHaveLength(1);
    expect(result.current.renderer.workerSurfacesActive).toBe(true);
    expect(result.current.renderer.status).toBe('ready');
    expect(result.current.renderer.layoutInWorker.isViewerSession?.(session)).toBe(true);
    if (workerProposals) {
      expect(registeredWorkerProposalAuthority(session)!.geometry()).not.toBeNull();
      expect(workerProposalAuthority(session)).toBe(registeredWorkerProposalAuthority(session));
    }
  } finally {
    io.unmount();
    unmount();
    project.mockRestore();
    release.mockRestore();
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
    const load = spyOn(full, 'loadState');
    registerRestore(() => load.mockRestore());
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
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
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
  'a combined worker open seeds only an unavailable editor worker with outcome=%s',
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
      if (outcome === 'throw') {
        await waitFor(() => expect(result.current.errors).toHaveLength(1));
        expect(result.current.errors[0]).toBeInstanceOf(DocxWorkerError);
        expect((result.current.errors[0] as DocxWorkerError).stage).toBe('open');
        expect(result.current.core.session).toBeNull();
        expect(result.current.core.previewing).toBe(false);
        expect(result.current.mainOpens).toEqual([]);
        unmount();
        return;
      }
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
      workerProposals: true, viewer: true, workerPreview: true,
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
    expect(workerOpenDocumentHeld(full)).toBe(true);
    options.oomStage = 'fontRequirements';
    await act(async () => {
      result.current.pipeline.runLayoutPipeline();
      await received('fontRequirements');
    });
    await waitFor(() => expect(result.current.errors).toHaveLength(1));
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

test.each([false, true])('a worker preview lays out before the full worker open with viewer=%s', async (viewer) => {
  const fullOpen = (request: ResidentEngineWorkerRequest) =>
    request.type === 'open' && request.previewBlocks === undefined;
  const { workers, posted, reply } = installWorker({ holdReply: fullOpen });
  const frames = holdFrames();
  const deferred = spyOn(replicaHelpers, 'deferWorkerOpenReplica');
  const main = viewer ? trackMainLoads() : null;
  try {
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, previewFirstPage: true, workerPreview: true, source: longBytes, viewer },
    });
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    expect(preview.isDisplayOnly()).toBe(true);
    if (viewer) expect(deferred).not.toHaveBeenCalled();
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
    if (viewer) {
      expect(workerOpenDocumentHeld(full)).toBe(true);
      expect(deferred).not.toHaveBeenCalled();
      expect(main!.loads.length).toBeGreaterThan(0);
      for (const load of main!.loads) expect(load.mock.calls).toHaveLength(0);
    }
    unmount();
  } finally {
    cleanup();
    deferred.mockRestore();
    main?.restore();
    frames.restore();
  }
});

test.each([false, true])('a shown viewer preview crash retries only the queued full document with terminal=%s', async (terminal) => {
  const fullOpen = (request: ResidentEngineWorkerRequest) =>
    request.type === 'open' && request.previewBlocks === undefined;
  const options: Parameters<typeof installWorker>[0] = { holdReply: fullOpen, holdRetryOpen: true };
  const { workers, posted } = installWorker(options);
  const frames = holdFrames();
  const main = trackMainLoads();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, previewFirstPage: true, workerPreview: true, source: longBytes, viewer: true,
    },
  });
  try {
    await waitFor(() => expect(result.current.core.previewing).toBe(true));
    const preview = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(preview));
    await waitFor(() => expect(posted.filter(fullOpen)).toHaveLength(1));
    const shown = result.current.renderer.displayList;
    expect(shown?.pages.length).toBeGreaterThan(0);
    options.holdReply = () => false;
    await act(async () => {
      workers[0].onerror?.({ message: 'preview worker crashed' } as ErrorEvent);
    });
    await waitFor(() => expect(posted.filter(fullOpen)).toHaveLength(2));
    expect(result.current.renderer.displayList).toBe(shown);
    act(() => result.current.presentFrame());
    act(() => frames.run());
    act(() => frames.run());
    if (terminal) {
      await act(async () => {
        workers[1].onerror?.({ message: 'full open retry crashed' } as ErrorEvent);
      });
      await waitFor(() => expect(result.current.errors).toHaveLength(1));
      const failure = result.current.errors[0];
      expect(failure).toBeInstanceOf(DocxWorkerError);
      expect((failure as DocxWorkerError).stage).toBe('open');
      expect(result.current.renderer.error).toBe(failure!);
      expect(result.current.core.session).toBeNull();
      for (let frame = 0; frame < 5; frame += 1) await act(async () => frames.run());
      expect(result.current.errors).toEqual([failure!]);
    } else {
      await act(async () => workers[1].release());
      await waitFor(() => expect(result.current.core.previewing).toBe(false));
      const full = result.current.core.session!;
      expect(workerOpenDocumentHeld(full)).toBe(true);
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
      expect(result.current.renderer.displayList?.pages.length).toBeGreaterThan(0);
      expect(result.current.errors).toEqual([]);
    }
    expect(workers).toHaveLength(2);
    expect(posted.filter(fullOpen)).toHaveLength(2);
    expect(posted.filter((request) => request.type === 'open' && request.previewBlocks !== undefined)).toHaveLength(1);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    for (const load of main.loads) expect(load.mock.calls).toHaveLength(0);
  } finally {
    unmount();
    main.restore();
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

test.each([false, true])('a failed worker open reports one typed error without a main seed with viewer=%s', async (viewer) => {
  const { workers, posted } = installWorker({ failOpen: true });
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, viewer },
  });
  await waitFor(() => expect(result.current.errors).toHaveLength(1));
  const failure = result.current.errors[0];
  expect(failure).toBeInstanceOf(DocxWorkerError);
  expect((failure as DocxWorkerError).stage).toBe('open');
  expect(result.current.renderer.error).toBe(failure);
  expect(result.current.core.session).toBeNull();
  expect(result.current.host).toBeNull();
  expect(result.current.mainOpens).toEqual([]);
  expect(workers).toHaveLength(2);
  expect(posted.filter((request) => request.type === 'open')).toHaveLength(2);
  expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  await act(async () => {});
  expect(result.current.errors).toEqual([failure]);
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

test.each(['capability', 'null'] as const)('a viewer with an unavailable worker (%s) reports a typed open error without a main seed', async (unavailable) => {
  globalThis.Worker = undefined as unknown as typeof Worker;
  const { result } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, viewer: true,
      ...(unavailable === 'null' ? { openInWorker: mock(async () => null) } : {}),
    },
  });
  await waitFor(() => expect(result.current.errors).toHaveLength(1));
  const failure = result.current.errors[0] as DocxWorkerError;
  expect(failure).toBeInstanceOf(DocxWorkerError);
  expect(failure.stage).toBe('open');
  expect(failure.cause).toEqual(new Error('The document worker is unavailable'));
  expect(result.current.host).toBeNull();
  expect(result.current.core.session).toBeNull();
  expect(result.current.mainOpens).toEqual([]);
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

test('eager hydration keeps worker rendering and proposal updates', async () => {
  const { workers, posted } = installWorker({ holdReply: () => false });
  const frames = holdFrames(true);
  let mirror: ReturnType<typeof spyOn<YrsSession, 'mirrorWorkerDocument'>> | undefined;
  let layoutHere: ReturnType<typeof spyOn<YrsSession, 'layoutDocumentWithRegionsRetainedJson'>> | undefined;
  let buildHere: ReturnType<typeof spyOn<YrsSession, 'buildDisplayListFrame'>> | undefined;
  let terminate: ReturnType<typeof spyOn<InProcessResidentWorker, 'terminate'>> | undefined;
  let session!: YrsSession;
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, workerProposals: true, allowHostProposals: true, styleResolver: true,
      onHostDocument: (next) => { session = next; },
      onLoad: () => {
        expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
        mirror = spyOn(session, 'mirrorWorkerDocument');
        layoutHere = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
        buildHere = spyOn(session, 'buildDisplayListFrame');
        terminate = spyOn(workers[0]!, 'terminate');
      },
    },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const openingApi = result.current.ref.current!;
    const request = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
    expect(openingApi.getEditorRef()).toBeNull();
    expectLoadingMutations(openingApi, result.current.pagedEditorRef.current!);
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(request).not.toHaveBeenCalled();
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    act(() => result.current.presentFrame());
    await frames.settleAndIdle(result.current.renderer.settledDisplayList(null, null, 'window'));
    act(() => {
      frames.run();
      frames.run();
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expectReadyMutations(openingApi, session, result.current.pagedEditorRef.current!, result.current.searchReveals);
    act(() => result.current.pipeline.scheduleLayout('remote', true));
    expect(openingApi.getEditorRef()?.getYrsSession()).toBe(session);
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([false]);
    expect(mirror).toHaveBeenCalledWith(null);
    expect(session.workerDocumentMirrored()).toBe(false);
    expect(terminate).not.toHaveBeenCalled();
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, 3000));
    expect(workers).toHaveLength(1);
    const bootstrap = posted.find((request) => request.type === 'bootstrap')!;
    expect(bootstrap.type).toBe('bootstrap');
    if (bootstrap.type !== 'bootstrap') throw new Error('Missing bootstrap');
    expect(bootstrap.opened).toBe(true);
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
    expect(workers[0]!.requests).toContain('sync');
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.renderer.workerSurfacesActive).toBe(true);
    expect(layoutHere).not.toHaveBeenCalled();
    expect(buildHere).not.toHaveBeenCalled();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    mirror?.mockRestore();
    layoutHere?.mockRestore();
    buildHere?.mockRestore();
    terminate?.mockRestore();
    frames.restore();
  }
});

test('a failed worker-recovery main open keeps rendering pinned to the main engine', async () => {
  const { workers, posted } = installWorker();
  const { result, unmount } = renderHook(useHarness, {
    initialProps,
  });
  let open: ReturnType<typeof spyOn<YrsSession, 'openDocx'>> | undefined;
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const api = result.current.ref.current!;
    const request = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
    expect(api.getEditorRef()).toBeNull();
    expectLoadingMutations(api, result.current.pagedEditorRef.current!);
    expect(request).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    const failure = new Error('Main open failed');
    open = spyOn(session, 'openDocx').mockImplementation(() => { throw failure; });
    const ready = awaitWorkerOpenReplica(session)!;
    act(() => expect(() => ensureWorkerOpenReplica(session)).toThrow(failure));
    await expect(ready).rejects.toBe(failure);
    expect(() => api.getEditorRef()).toThrow(failure);
    expect(() => result.current.renderer.layoutInWorker(session, '{}')).toThrow(failure);
    expect(workers).toHaveLength(1);
    expect(posted.some((request) => request.type === 'bootstrap')).toBe(false);
    expect(result.current.errors).toEqual([failure]);
  } finally {
    open?.mockRestore();
    unmount();
  }
});

test('a successful worker-recovery main open permits pending reads, navigation and mutations', async () => {
  const { workers, posted } = installWorker();
  const frames = holdFrames(true);
  const source = await longFixture(2);
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, source },
  });
  try {
    await frames.waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const api = result.current.ref.current!;
    const editor = result.current.pagedEditorRef.current!;
    const request = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
    const ensure = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
    expect(api.getEditorRef()).toBeNull();
    expectLoadingMutations(api, editor);
    expect(request).not.toHaveBeenCalled();
    expect(ensure).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([]);
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    const ready = awaitWorkerOpenReplica(session)!;
    act(() => ensureWorkerOpenReplica(session));
    await frames.untilCommitted(ready);
    expect(result.current.core.replicaReady).toBe(true);
    expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(false);
    expect(ensure.mock.calls).toEqual([[session], [session]]);
    expect(api.findInDocument('paragraph')).toContainEqual({
      paraId: 'body:p0', match: 'paragraph', before: 'First ', after: '',
    });
    expectReadyMutations(api, session, editor, result.current.searchReveals);
    expect(result.current.mainOpens).toEqual([true]);
    expect(request).not.toHaveBeenCalled();
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(workers).toHaveLength(1);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    frames.restore();
  }
});

test('opening the comments sidebar keeps a viewer document held in the worker', async () => {
  const { posted } = installWorker();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, viewer: true, readOnly: true },
  });
  try {
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    const session = result.current.core.session!;
    act(() => result.current.openCommentsSidebar());
    await act(async () => {});
    expect(workerOpenDocumentHeld(session)).toBe(true);
    expect(result.current.core.replicaReady).toBe(false);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.renderer.workerSurfacesActive).toBe(true);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
  }
});

test('a first-layout font setup failure starts the replica for pending reads and commands', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const session = result.current.core.session!;
  const failure = new Error('font setup failed');
  const registerFont = spyOn(session, 'registerFont').mockImplementation(() => { throw failure; });
  try {
    const calls = Promise.allSettled([
      result.current.ref.current!.readParagraphs({ view: 'accepted' }),
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
    expect(settled[1]).toEqual({ status: 'fulfilled', value: true });
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.renderer.frame).toBeNull();
  } finally {
    registerFont.mockRestore();
  }
});

test('a worker open without a frame or error starts the replica after the bounded wait', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const peerFallback = holdPeerFallback();
  const layoutFallback = holdLayoutFallback();
  const { result } = renderHook(useHarness, { initialProps });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(await result.current.ref.current!.save()).toEqual(new ArrayBuffer(0));
  expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  expect(result.current.mainOpens).toEqual([]);
  const calls = Promise.allSettled([
    result.current.ref.current!.readParagraphs({ view: 'accepted' }),
  ]);
  const completed = { value: false };
  void calls.then(() => { completed.value = true; });
  expect(posted.map((request) => request.type)).toEqual(['open']);
  act(() => peerFallback.advance(10_000));
  await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true), { timeout: 7000 });
  expect(result.current.renderer.frame).toBeNull();
  expect(result.current.errors).toEqual([]);
  expect(result.current.mainOpens).toEqual([]);
  expect(completed.value).toBe(false);
  await act(async () => { workers[0].release(); });
  await waitFor(() => expect(layoutFallback.timers.size).toBe(2));
  act(() => layoutFallback.advance(3000));
  await waitFor(() => expect(completed.value).toBe(true));
  const settled = await calls;
  expect(settled[0]).toMatchObject({ status: 'fulfilled', value: { ok: true } });
  expect(result.current.mainOpens).toEqual([false]);
  expect(result.current.core.replicaReady).toBe(true);
}, 15_000);

test('a failed worker open reports its typed error without parsing a main fallback', async () => {
  const invalid = Uint8Array.of(1, 2, 3);
  const { workers, posted } = installWorker({ failOpen: true });
  const { result } = renderHook(useHarness, { initialProps: { ...initialProps, source: invalid } });
  await waitFor(() => expect(result.current.errors).toHaveLength(1));
  const failure = result.current.errors[0] as DocxWorkerError;
  expect(failure).toBeInstanceOf(DocxWorkerError);
  expect(failure.stage).toBe('open');
  expect(failure.cause).toEqual(new Error('open failed'));
  expect(result.current.renderer.error).toBe(failure);
  expect(workers).toHaveLength(2);
  expect(posted.filter((request) => request.type === 'open')).toHaveLength(2);
  expect(result.current.mainOpens).toEqual([]);
  expect(result.current.core.session).toBeNull();
});

test('a shared collaboration update keeps the existing join path', async () => {
  const shared = await createYrsSession();
  sessions.push(shared);
  shared.openDocx(bytes, true);
  const { workers, posted } = installWorker();
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, experimentalWorkerOpen: false, collaboration: { initialUpdate: shared.encodeState() } },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  expect(workers).toHaveLength(0);
  expect(result.current.core.replicaReady).toBe(true);
  expect(texts(result.current.core.session!)).toEqual(texts(shared));

  const source = await longFixture(3);
  const sharedEditor = await createYrsSession();
  sessions.push(sharedEditor);
  sharedEditor.openDocx(source, true);
  const paragraph = sharedEditor.paragraphs('body')[0]!;
  sharedEditor.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Joined ');
  const initialUpdate = sharedEditor.encodeState();
  const displayList = await import('./useDisplayList');
  const useRenderer = displayList.useCanvasRenderer;
  const workerModes: boolean[] = [];
  const choice = spyOn(displayList, 'useCanvasRenderer').mockImplementation((...args) => {
    workerModes.push(args[5] === true);
    return useRenderer(...args);
  });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const ref = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: { addEventListener() {}, removeEventListener() {}, ready: Promise.resolve() },
  });
  try {
    const view = render(
      <DocxEditor
        ref={ref}
        documentBuffer={source.slice().buffer as ArrayBuffer}
        collaboration={{ initialUpdate }}
        onError={(error) => errors.push(error)}
      />
    );
    await waitFor(() => expect(view.container.querySelector('.canvas-page')).not.toBeNull(), {
      timeout: 20_000,
    });
    const joined = await ref.current!.readParagraphs({ view: 'accepted' });
    expect(joined).toMatchObject({ ok: true });
    if (!joined.ok) throw new Error('The joined paragraphs are unavailable');
    expect(joined.paragraphs.map(({ text }) => text)).toEqual(texts(sharedEditor).body);
    expect(joined.paragraphs[0]!.text).toBe('Joined First paragraph');
    expect(workerModes.length).toBeGreaterThan(0);
    expect(workerModes.every((worker) => !worker)).toBe(true);
    expect(posted.filter((request) => request.type === 'open')).toEqual([]);
    expect(warn.mock.calls.some(([message]) => String(message).includes('collaboration.initialUpdate requires the in-thread engine'))).toBe(true);
    expect(errors).toEqual([]);
  } finally {
    choice.mockRestore();
    warn.mockRestore();
  }
}, 40_000);

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
  test(`terminal OOM during ${stage} rejects pending reads and commands`, async () => {
    const { workers, posted } = installWorker({ oomStage: stage });
    const { result } = renderHook(useHarness, { initialProps });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const calls = Promise.allSettled([
      result.current.ref.current!.readParagraphs({ view: 'accepted' }),
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
    expect(settled).toEqual(Array.from({ length: 3 }, () => ({ status: 'rejected', reason: failure })));
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
  viewer: true,
  workerProposals: true,
  allowHostProposals: true,
};

async function openWorkerProposals(props: HarnessProps = workerProposalProps) {
  const harness = renderHook(useHarness, { initialProps: props });
  await waitFor(() => expect({ workerProposalsReady: harness.result.current.core.workerProposalsReady })
    .toEqual({ workerProposalsReady: true }));
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
      expect(failure).toBeInstanceOf(DocxWorkerError);
      expect((failure as DocxWorkerError).stage).toBe('layout');
      expect(failure?.cause).toEqual(new Error(path === 'no snapshot'
        ? 'Resident worker snapshot was not available'
        : 'The document worker cannot lay out this viewer document'));
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

async function failIdleWorker(worker: InProcessResidentWorker, message: string): Promise<void> {
  await act(async () => {
    worker.onerror?.({ message } as ErrorEvent);
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

test('a worker that fails while idle after load is replaced once and keeps painting and reading', async () => {
  const { workers, posted } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  try {
    expect(workers).toHaveLength(1);
    await failIdleWorker(workers[0], 'worker lost');
    await waitFor(() => expect(workers).toHaveLength(2));
    await waitFor(() => expect(posted.filter((request) => request.type === 'bootstrap')).toHaveLength(2));
    await waitFor(() => expect({ workerSurfacesActive: result.current.renderer.workerSurfacesActive })
      .toEqual({ workerSurfacesActive: true }));
    await waitFor(() => expect({ queriesReady: result.current.renderer.queries?.isReady() })
      .toEqual({ queriesReady: true }));
    expect(result.current.renderer.frame).not.toBeNull();
    expect(result.current.renderer.queries!.pageCount()).toBeGreaterThan(0);
    const read = await result.current.ref.current!.readParagraphs({ view: 'accepted' });
    expect(read).toMatchObject({ ok: true });
    if (!read.ok) throw new Error(read.failure.message);
    expect(read.paragraphs).toHaveLength(205);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(2);
    const api = result.current.ref.current!;
    const paragraph = (await api.getParagraphIdentities()).paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const proposals = await api.getProposals();
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: proposals.version,
        proposals: [{
          id: 'after-recovery', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Recovered ',
        }],
      })).toMatchObject({ ok: true });
    });
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.renderer.error).toBeNull();
    expect(result.current.errors).toEqual([]);
    expect(errorLog).not.toHaveBeenCalled();
    expect(warning.mock.calls.map(([message]) => message)).toContain(
      '[CanvasRenderer] Resident engine worker failed; starting a fresh worker'
    );
  } finally {
    unmount();
    warning.mockRestore();
    errorLog.mockRestore();
  }
}, 15_000);

test('a replacement worker that fails while idle too fails the document once', async () => {
  const { workers } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  try {
    await failIdleWorker(workers[0], 'worker lost');
    await waitFor(() => expect(result.current.renderer.workerSurfacesActive).toBe(true));
    expect(workers).toHaveLength(2);
    await failIdleWorker(workers[1], 'worker lost again');
    await waitFor(() => expect(result.current.renderer.error).not.toBeNull());
    const failure = result.current.renderer.error!;
    expect(failure.message).toContain('worker lost again');
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(workers).toHaveLength(2);
    expect(errorLog.mock.calls.filter(([, error]) => error === failure)).toEqual([
      ['[CanvasRenderer] Resident engine worker failed again', failure],
    ]);
    expect(result.current.errors).toEqual([failure]);
    expect(result.current.mainOpens).toEqual([]);
  } finally {
    unmount();
    warning.mockRestore();
    errorLog.mockRestore();
  }
}, 15_000);

test('a worker holding committed proposals that fails while idle fails the document without a fresh worker', async () => {
  const { workers, posted } = installWorker();
  const { result, unmount } = await openWorkerProposals();
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
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
    await waitFor(() => expect(sourceVersionOf(result.current.renderer.queries)).toBe(held.version));
    await failIdleWorker(workers[0], 'worker lost');
    await waitFor(() => expect(result.current.renderer.error).not.toBeNull());
    const failure = result.current.renderer.error!;
    expect(failure.message).toContain('worker lost');
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    await expect(api.getProposals()).rejects.toBe(failure);
    expect(workers).toHaveLength(1);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([failure]);
    expect(errorLog.mock.calls.filter(([, error]) => error === failure)).toEqual([
      ['[CanvasRenderer] Resident engine worker holding proposals failed', failure],
    ]);
    expect(warning.mock.calls.some(([message]) => String(message).includes('starting a fresh worker'))).toBe(false);
  } finally {
    unmount();
    warning.mockRestore();
    errorLog.mockRestore();
  }
}, 15_000);

test('a worker that fails while a layout request waits keeps the request path routing', async () => {
  const { workers, posted, received } = installWorker({ holdReply: (request) => request.type === 'bootstrap' });
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  const { result, unmount } = renderHook(useHarness, { initialProps });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    act(() => result.current.pipeline.runLayoutPipeline());
    await act(async () => { await received('bootstrap'); });
    await failIdleWorker(workers[0], 'worker lost mid-load');
    await waitFor(() => expect(result.current.renderer.frame).not.toBeNull());
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
    expect(workers).toHaveLength(1);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(errorLog.mock.calls.map(([message]) => message)).toContain(
      '[CanvasRenderer] Resident engine worker unavailable; laying out on the main thread'
    );
    expect(warning.mock.calls.some(([message]) => String(message).includes('starting a fresh worker'))).toBe(false);
    expect(result.current.renderer.error).toBeNull();
  } finally {
    unmount();
    warning.mockRestore();
    errorLog.mockRestore();
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
  'viewer revisions are asked after worker completion is queued without loading a replica, with onWorkerRevisions=%s',
  async (withCallback) => {
    const source = await withTimeout(longFixture(1200), 5000, 'viewer fixture');
    const bounds = { timeout: 1000 };
    let completionReplied = false;
    const asked: boolean[] = [];
    const onWorkerRevisions = mock(() => {});
    const { workers, posted, received } = installWorker({
      holdState: true,
      holdBootstrap: true,
      holdCompletion: true,
      holdReply: () => false,
      revisionCount: 1,
      onRevisionCount: () => asked.push(completionReplied),
    });
    const frames = holdFrames();
    const props = {
      ...workerProposalProps,
      source,
      followCompletion: true,
      onWorkerRevisions: withCallback ? onWorkerRevisions : undefined,
    };
    const { result, unmount } = renderHook(useHarness, { initialProps: props });
    try {
      await withTimeout(act(async () => { await received('open', 0, 5000); }), 6000, 'open React updates');
      await waitFor(() => expect(result.current.host).not.toBeNull(), bounds);
      const session = result.current.core.session!;
      const receive = workers[0].onmessage;
      workers[0].onmessage = (event) => {
        if (posted.some((request) =>
          request.type === 'completeLayout' && request.id === event.data.id
        )) completionReplied = true;
        receive?.(event);
      };
      await waitFor(() => expect(posted.map((request) => request.type)).toContain('bootstrap'), bounds);
      expect(result.current.core.workerProposalsReady).toBe(false);
      expect(posted.some((request) => request.type === 'proposal')).toBe(false);
      await withTimeout(act(async () => { workers[0].release(); }), 1000, 'bootstrap React updates');
      await waitFor(() => expect(posted.map((request) => request.type)).toContain('completeLayout'), {
        timeout: 5000,
      });
      await waitFor(() => expect(result.current.renderer.pendingCompletion).toBeNull(), bounds);
      await withTimeout(act(async () => {}), 1000, 'completion React updates');
      act(() => result.current.presentFrame());
      act(() => frames.run());
      act(() => frames.run());
      await waitFor(() => expect(posted.filter((request) =>
        request.type === 'revisionCount'
      )).toHaveLength(1), bounds);
      expect(asked).toEqual([false]);
      expect(completionReplied).toBe(false);
      expect(posted.findIndex((request) => request.type === 'revisionCount')).toBeGreaterThan(
        posted.findIndex((request) => request.type === 'completeLayout')
      );
      if (withCallback) {
        await waitFor(() => expect(onWorkerRevisions).toHaveBeenCalledTimes(1), bounds);
      }
      act(() => result.current.presentFrame());
      act(() => frames.run());
      act(() => frames.run());
      await withTimeout(act(async () => {}), 1000, 'revision React updates');
      expect(onWorkerRevisions).toHaveBeenCalledTimes(withCallback ? 1 : 0);
      expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
      expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
      expect(result.current.core.replicaReady).toBe(false);
      expect(replicaHelpers.workerOpenReplicaPending(session)).toBe(true);
      expect(session.storyIds()).toEqual([]);
      expect(result.current.mainOpens).toEqual([]);
      await waitFor(() => expect(result.current.core.workerProposalsReady).toBe(true), bounds);
      expect(posted.filter((request) => request.type === 'proposal')).toHaveLength(1);
      expect(completionReplied).toBe(false);
      await withTimeout(act(async () => { workers[0].release(); }), 1000, 'released completion React updates');
      await waitFor(() => expect(result.current.core.workerProposalsReady).toBe(true), bounds);
      expect(posted.filter((request) => request.type === 'proposal')).toHaveLength(1);
      expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
      expect(result.current.core.replicaReady).toBe(false);
      expect(workerOpenDocumentHeld(session)).toBe(true);
      expect(result.current.mainOpens).toEqual([]);
      expect(result.current.errors).toEqual([]);
    } finally {
      try {
        unmount();
      } finally {
        frames.restore();
      }
    }
  },
  30_000
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
  const { result, rerender, unmount } = renderHook(useHarness, {
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
    act(() => rerender({ ...workerProposalProps, viewer: false }));
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

test.each([false, true])('a sync ref call during the first in-flight proposal keeps the worker\'s proposals with viewer=%s', async (viewer) => {
  const { workers, posted, received, reply } = installWorker({
    holdReply: (request) => request.type === 'proposal' && request.operation.kind === 'propose',
  });
  const frames = holdFrames();
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...workerProposalProps, viewer, styleResolver: !viewer },
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
      if (!viewer) {
        const request = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
        expect(api().getDocument()).toBeNull();
        expectLoadingMutations(api(), result.current.pagedEditorRef.current!, paragraph.paraId);
        expect(request).not.toHaveBeenCalled();
        expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
        expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
      } else
      act(() => { expect(() => api().getDocument()).toThrow(viewer ? DocxAsyncOnlyError : DocxReplicaNotReadyError); });
      if (viewer) {
        expect(() => api().getPageContent(1)).toThrow(DocxAsyncOnlyError);
        expect(() => api().findInDocument('paragraph')).toThrow(DocxAsyncOnlyError);
        expect(api().getSelectionInfo()).toBeNull();
      }
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
        if (!viewer) {
          const ready = api().flushPendingInput();
          await requestWorkerOpenReplica(session);
          await ready;
        }
      });
      if (!viewer) expectReadyMutations(api(), session, result.current.pagedEditorRef.current!, result.current.searchReveals, paragraph.paraId);
      if (!viewer) expect(api().getDocument()).not.toBeNull();
      expect(result.current.core.replicaReady).toBe(!viewer);
      expect(workerOpenDocumentHeld(session)).toBe(viewer);
      expect(result.current.mainOpens).toEqual(viewer ? [] : [false]);
      expect(session.workerDocumentMirrored()).toBe(viewer);
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
    initialProps: { ...workerProposalProps, viewer: false, styleResolver: true },
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
      const request = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
      expect(api().getDocument()).toBeNull();
      expectLoadingMutations(api(), result.current.pagedEditorRef.current!, paragraph.paraId);
      expect(result.current.mainOpens).toEqual([]);
      expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
      expect(request).not.toHaveBeenCalled();
      expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
      act(() => { ensureWorkerOpenReplica(session); });
      expectReadyMutations(api(), session, result.current.pagedEditorRef.current!, result.current.searchReveals, paragraph.paraId);
      expect(result.current.core.documentFromYrs()).not.toBeNull();
      expect(api().getDocument()).not.toBeNull();
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

test('a refused viewer proposal keeps its live worker and held document', async () => {
  const { workers, posted } = installWorker({ failProposal: true });
  const { result, unmount } = await openWorkerProposals({ ...workerProposalProps, source: bytes });
  try {
    const api = result.current.ref.current!;
    const session = result.current.core.session!;
    const identities = await api.getParagraphIdentities();
    const paragraph = identities.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
    const initial = await api.getProposals();
    await act(async () => {
      await expect(api.proposeChanges({
        expectVersion: initial.version,
        proposals: [{
          id: 'refused-proposal', paragraph,
          suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
          op: 'insertText', at: 'start', text: 'Refused ',
        }],
      })).rejects.toThrow('proposal failed');
    });
    expect(await api.getProposals()).toEqual(initial);
    expect(workerOpenDocumentHeld(session)).toBe(true);
    expect(workerProposalAuthority(session)!.holdsWorkerState()).toBe(false);
    expect(workers).toHaveLength(1);
    expect(posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.renderer.error).toBeNull();
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
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
  const { result, rerender, unmount } = await openWorkerProposals();
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
    act(() => rerender({ ...workerProposalProps, viewer: false }));
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
  const { result, rerender, unmount } = await openWorkerProposals();
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
    act(() => rerender({ ...workerProposalProps, viewer: false }));
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

test('without worker proposals the first editor round waits for hydration and then uses the worker', async () => {
  const { posted, workers } = installWorker({ holdState: true });
  const { result } = renderHook(useHarness, {
    initialProps: { ...initialProps, workerProposals: false, allowHostProposals: true },
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  const session = result.current.core.session!;
  expect(registeredWorkerProposalAuthority(session)).toBeNull();
  const peerPropose = spyOn(session, 'proposeChanges');
  let settled = false;
  const call = result.current.ref.current!.proposeChanges({ expectVersion: session.version(), proposals: [] });
  void call.then(() => { settled = true; });
  expect(registeredWorkerProposalAuthority(session)).toBeNull();
  expect(posted.some((request) => request.type === 'proposal')).toBe(false);
  expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
  expect(result.current.mainOpens).toEqual([]);
  expect(settled).toBe(false);
  act(() => { requestWorkerOpenReplica(session); result.current.pipeline.runLayoutPipeline(); });
  await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
  expect(settled).toBe(false);
  await act(async () => { workers[0].release(); expect(await call).toMatchObject({ ok: true }); });
  expect(settled).toBe(true);
  expect(result.current.mainOpens).toEqual([false]);
  expect(peerPropose).not.toHaveBeenCalled();
  expect(posted.filter((request) => request.type === 'proposal').at(-1)?.operation.kind).toBe('propose');
});

test('a first editor round uses the base peer call when pending hydration falls back to source', async () => {
  const { posted } = installWorker({ failState: true });
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, workerProposals: false, allowHostProposals: true },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const peerPropose = spyOn(session, 'proposeChanges');
    const round = result.current.ref.current!.proposeChanges({ expectVersion: session.version(), proposals: [] });
    expect(registeredWorkerProposalAuthority(session)).toBeNull();
    expect(posted.some((request) => request.type === 'proposal')).toBe(false);
    await act(async () => {
      await requestWorkerOpenReplica(session);
      expect(await round).toMatchObject({ ok: true, snapshot: { proposals: [] } });
    });
    expect(peerPropose).toHaveBeenCalledTimes(1);
    expect(result.current.mainOpens).toEqual([true]);
    expect(result.current.core.replicaReady).toBe(true);
    expect(registeredWorkerProposalAuthority(session)).toBeNull();
    expect(posted.some((request) => request.type === 'proposal')).toBe(false);
  } finally { unmount(); }
});

test('a font preflight OOM after a proposal mirrors retires without failing peer rounds or saves', async () => {
  let holdMutation = false;
  let holdRequirements = false;
  const options: Parameters<typeof installWorker>[0] = {
    holdReply: (request) => (holdMutation && request.type === 'proposal' && request.operation.kind === 'propose') ||
      (holdRequirements && request.type === 'fontRequirements'),
  };
  const { posted, workers, received, reply, responses } = installWorker(options);
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, workerProposals: false, allowHostProposals: true },
  });
  const errorLog = spyOn(console, 'error').mockImplementation(() => {});
  let frames: ReturnType<typeof holdFrames> | undefined;
  let unmountPreflight: (() => void) | undefined;
  let unmountIO: (() => void) | undefined;
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    await act(async () => { await requestWorkerOpenReplica(session); });
    const api = result.current.ref.current!;
    await act(async () => {
      expect(await api.proposeChanges({ expectVersion: session.version(), proposals: [] })).toMatchObject({ ok: true });
      await api.whenLayoutComplete({ timeoutMs: 5000 });
    });
    const authority = registeredWorkerProposalAuthority(session)!;
    expect(authority.initialized).toBe(true);
    const paragraph = (await api.getParagraphIdentities()).paragraphs.find((entry) =>
      entry.session?.story === 'body'
    )!.session!;
    const preflightErrors: Array<[Error, YrsSession]> = [];
    const fontId = session.registerFont(font);
    const preflight = renderHook(() => useLayoutPipeline({
      document: result.current.host!.document, session, renderEnv: {}, pageGap: 24, zoom: 1,
      residentMeasurementConfig: (requirements) => ({
        fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [fontId]])),
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: { noLeading: false, doNotExpandShiftReturn: false }, authoritativeShaping: true,
      }),
      deferLayoutPass: () => false,
      pagesContainerRef: { current: null }, viewportLayoutRef: { current: null },
      syncCoordinator: new LayoutSelectionGate(), getScrollContainer: () => null,
      experimentalWorkerOpen: true,
      fontRequirementsInWorker: result.current.renderer.fontRequirementsInWorker,
      layoutInWorker: result.current.renderer.layoutInWorker,
      onError: (error, owner) => preflightErrors.push([error, owner]),
    }));
    unmountPreflight = preflight.unmount;
    frames = holdFrames();
    holdMutation = true;
    const previous = posted.at(-1)!.id;
    const round = api.proposeChanges({
      expectVersion: session.version(),
      proposals: [{
        id: 'before-font-oom', paragraph,
        suggest: { author: 'Host', date: '2026-10-06T00:00:00Z' },
        op: 'insertText', at: 'start', text: 'Mirrored ',
      }],
    });
    void round.catch(() => {});
    const mutation = await received('proposal', previous);
    expect(authority.holdsWorkerState()).toBe(true);
    holdRequirements = true;
    options.oomStage = 'fontRequirements';
    act(() => preflight.result.current.runLayoutPipeline());
    const requirements = await received('fontRequirements', mutation.id);
    expect(responses.get(requirements)).toMatchObject({ ok: false, outOfMemory: true });
    await act(async () => { reply(mutation); expect(await round).toMatchObject({ ok: true }); });
    expect(session.paragraphs('body')[0]!.text).toContain('Mirrored ');
    expect(authority.holdsWorkerState()).toBe(false);
    expect(authority.retirementReason()).toBeNull();
    const peerRequirements = spyOn(session, 'layoutFontRequirementsJson');
    const peerLayout = spyOn(session, 'layoutDocumentWithRegionsRetainedJson');
    await act(async () => { reply(requirements); });
    expect(authority.retirementReason()).toBe('source-fallback');
    expect(workerProposalFailure(session)).toBeUndefined();
    expect(preflightErrors).toHaveLength(1);
    expect(preflightErrors[0]![0]).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(preflightErrors[0]![1]).toBe(session);
    await act(async () => { await frames!.until(result.current.renderer.settledDisplayList(null, null)); });
    expect(peerRequirements).toHaveBeenCalled();
    expect(peerLayout).toHaveBeenCalled();
    expect(result.current.renderer.workerSurfacesActive).toBe(false);
    expect(result.current.renderer.status).toBe('ready');
    expect(sourceVersionOf(result.current.renderer.queries)).toBe(session.version());
    const workerRounds = posted.filter((request) => request.type === 'proposal').length;
    await act(async () => {
      expect(await api.proposeChanges({
        expectVersion: session.version(),
        proposals: [{
          id: 'after-font-oom', paragraph,
          suggest: { author: 'Host', date: '2026-10-06T00:00:00Z' },
          op: 'insertText', at: 'end', text: 'Recovered ',
        }],
      })).toMatchObject({ ok: true });
    });
    expect(authority.snapshot()!.proposals.map(({ id }) => id)).toEqual(['before-font-oom', 'after-font-oom']);
    expect(session.getProposals().proposals).toEqual([]);
    expect(session.paragraphs('body')[0]!.text).toBe('Mirrored Page mapRecovered ');
    expect(posted.filter((request) => request.type === 'proposal')).toHaveLength(workerRounds);
    const saved: ArrayBuffer[] = [];
    const saveErrors: Error[] = [];
    const io = renderHook(() => useFileIO({
      pagedEditorRef: result.current.pagedEditorRef, viewerSession: false,
      resolveImage: () => null, comments: [], documentName: undefined,
      onSave: (buffer) => saved.push(buffer), downloadOnSave: false,
      onError: (error) => saveErrors.push(error),
      onOpen: undefined, onPrint: undefined, onDocumentNameChange: undefined,
      loadBuffer: async () => {}, focusActiveEditor: () => {},
    }));
    unmountIO = io.unmount;
    let buffer: ArrayBuffer | null = null;
    await act(async () => { buffer = await io.result.current.handleSave(); });
    expect(buffer).toBeInstanceOf(ArrayBuffer);
    expect(saved).toEqual([buffer!]);
    expect(saveErrors).toEqual([]);
    expect(workerProposalFailure(session)).toBeUndefined();
    expect(workers).toHaveLength(1);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmountIO?.();
    unmountPreflight?.();
    unmount();
    frames?.restore();
    errorLog.mockRestore();
  }
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
  const { result, rerender } = renderHook(useHarness, {
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
  act(() => rerender({ ...workerProposalProps, viewer: false }));
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
  const { result, rerender } = renderHook(useHarness, {
    initialProps: workerProposalProps,
  });
  await waitFor(() => expect(result.current.host).not.toBeNull());
  await act(async () => { await received('bootstrap'); });
  const session = result.current.core.session!;
  await act(async () => { await result.current.ref.current!.getProposals(); });
  expect(session.workerDocumentMirrored()).toBe(true);
  act(() => rerender({ ...workerProposalProps, viewer: false }));
  await act(async () => { await requestWorkerOpenReplica(session); });
  expect(result.current.mainOpens).toEqual([true]);
  expect(session.workerDocumentMirrored()).toBe(false);
  expect(await result.current.ref.current!.setProposalStates({
    expectVersion: session.version(), expectPreviewVersion: session.getProposals().previewVersion,
    changes: [],
  })).toMatchObject({ ok: true });
});

async function openingEditor(workerPreview = true, viewer = false, options: {
  delayQueries?: boolean;
  source?: Uint8Array;
  residentInput?: boolean;
  publishLayout?: boolean;
  holdFullLayout?: boolean;
  holdReply?: (request: ResidentEngineWorkerRequest) => boolean;
  onFirstPagePainted?: (input: { click(position: number): void; type(text: string): void }) => void;
} = {}) {
  const isFullOpen = (request: ResidentEngineWorkerRequest) =>
    request.type === 'open' && request.previewBlocks === undefined;
  let layoutHeld = false;
  let fullOpenId = 0;
  const holdReply = options.holdFullLayout
    ? (request: ResidentEngineWorkerRequest) => options.holdReply?.(request) === true ||
      (layoutHeld && (request.type === 'sync' || request.type === 'completeLayout'))
    : options.holdReply;
  const worker = installWorker(workerPreview
    ? { holdReply: (request) => isFullOpen(request) || holdReply?.(request) === true, holdState: true }
    : { holdOpen: true, holdState: true, holdReply });
  const frames = holdFrames(true);
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const editor = createRef<PagedEditorRef>();
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  let releaseQueries = () => {};
  const queryReady = new Promise<void>((resolve) => { releaseQueries = resolve; });
  let queriesReleased = !options.delayQueries;
  let previousInputQueries: DisplayListQueries | null = null;
  const inputQueryFacades = new WeakMap<DisplayListQueries, DisplayListQueries>();
  const inputQueries = (queries: DisplayListQueries | null | undefined) => {
    if (!queries || !options.delayQueries) return queries;
    let facade = inputQueryFacades.get(queries);
    if (!facade) {
      const loaded = spyOn(queryEngines, 'loadedRustDisplayListQueryEngine').mockReturnValue(null);
      const loading = spyOn(queryEngines, 'loadRustDisplayListQueryEngine').mockImplementation(async () => {
        await queryReady;
        await queries.whenReady();
        const engine: queryEngines.RustDisplayListQueryEngine = {
          hitTestRegionsJson: (_list, pageIndex, x, y) => JSON.stringify(queries.hitTestRegions(pageIndex, x, y)),
          rangeRectsJson: () => '[]',
        };
        return engine;
      });
      try {
        facade = createDisplayListQueries(queries.displayList, undefined, previousInputQueries, harness.core.session);
      } finally {
        loaded.mockRestore();
        loading.mockRestore();
      }
      previousInputQueries = facade;
      inputQueryFacades.set(queries, facade);
    }
    return facade;
  };
  const props = { ...initialProps, source: options.source ?? (options.holdFullLayout ? await longFixture(1200) : longBytes),
    previewFirstPage: true, workerPreview,
    viewer, workerProposals: viewer, readOnly: viewer };
  const selections: Array<ReturnType<YrsSession['cellSelection']>> = [];
  function Editable({ source, generation, readOnly = viewer, viewer: viewerRequested = viewer }: Pick<HarnessProps, 'source' | 'generation' | 'readOnly' | 'viewer'>) {
    const viewerSession = useViewerSession(props.experimentalWorkerOpen, viewerRequested, generation);
    harness = useHarness({
      ...props, source, generation, readOnly, viewer: viewerSession,
      handleSave: async () => {
        const current = editor.current;
        const session = current?.getYrsSession();
        if (session && replicaHelpers.workerOpenReplicaStarted(session)) {
          await awaitWorkerOpenReplica(session);
        }
        await current?.flushPendingInput();
        return new ArrayBuffer(0);
      },
    });
    const layoutInWorker = useMemo<LayoutInWorker>(() => {
      const layout = harness.renderer.layoutInWorker;
      if (!options.holdFullLayout) return layout;
      let pending: ReturnType<LayoutInWorker> = null;
      return Object.assign((session: YrsSession, request: string) => {
        if (!layoutHeld || session.isDisplayOnly() || !replicaHelpers.workerOpenReplicaPending(session)) {
          return layout(session, request);
        }
        return pending ??= layout(session, request);
      }, { prewarm: layout.prewarm, ownsDocument: layout.ownsDocument });
    }, [harness.core.session, harness.renderer.layoutInWorker]);
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        readOnly={readOnly || harness.core.opening} holdInput={!readOnly && harness.core.opening} inputScope={generation}
        viewerDocumentRead={viewerSession ? harness.renderer.readWorkerDocument : undefined}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={layoutInWorker}
        onLayoutComputed={options.publishLayout || (options.holdFullLayout && !harness.core.previewing)
          ? harness.renderer.onLayoutComputed : undefined}
        applyResidentInput={options.residentInput ? harness.renderer.applyInput : undefined}
        canvasHostRef={canvasHost} displayListQueries={queriesReleased ? harness.renderer.queries : null}
        inputQueries={viewerSession ? undefined : inputQueries(harness.renderer.inputQueries)}
        onYrsSelectionChange={() => selections.push(harness.core.session!.cellSelection())} />
    </>;
  }
  const view = render(<Editable {...props} />);
  const pointAt = (position: number) => {
    const queries = harness.renderer.queries!;
    const caret = queries.caretRect(position)!;
    const size = queries.pageSize(0)!;
    const canvas = canvasHost.current!.firstElementChild!;
    canvas.getBoundingClientRect = () => ({ left: 0, top: 0, right: size.width,
      bottom: size.height, ...size }) as DOMRect;
    return { clientX: caret.x, clientY: caret.y + caret.height / 2, button: 0, detail: 1 };
  };
  const mouseDown = (position: number) => {
    const point = pointAt(position);
    const canvas = canvasHost.current!.firstElementChild!;
    if (fireEvent.mouseDown(canvas, point)) (document.activeElement as HTMLElement | null)?.blur();
  };
  const mouseMove = (position: number) => fireEvent.mouseMove(canvasHost.current!.firstElementChild!, pointAt(position));
  const mouseUp = (position: number) => fireEvent.mouseUp(window, pointAt(position));
  const click = (position: number) => {
    const point = pointAt(position);
    const canvas = canvasHost.current!.firstElementChild!;
    mouseDown(position);
    fireEvent.mouseUp(window, point);
    fireEvent.click(canvas, point);
  };
  const type = (text: string) => {
    for (const key of text) {
      const target = document.activeElement!;
      fireEvent.keyDown(target, { key });
      fireEvent.input(target, { target: { value: key } });
    }
  };
  const waitForPreview = async (generation = 1) => {
    await waitFor(() => {
      expect(harness.core.previewing).toBe(true);
      expect(harness.core.sessionGeneration).toBe(generation);
    });
    const preview = harness.core.session!;
    act(() => harness.pipeline.runLayoutPipeline());
    await waitFor(() => expect(harness.renderer.presentedEngine).toBe(preview));
    await waitFor(() => expect(harness.renderer.queries?.isReady()).toBe(true));
    act(() => {
      harness.presentFrame();
      options.onFirstPagePainted?.({ click, type });
    });
    return preview;
  };
  try {
    const preview = await waitForPreview();
    await waitFor(() => expect(worker.posted.some(isFullOpen)).toBe(true));
    const switchToFull = async () => {
      act(() => frames.run());
      act(() => frames.run());
      await waitFor(() => expect(worker.posted.some(isFullOpen)).toBe(true));
      const request = worker.posted.filter(isFullOpen).at(-1)!;
      if (workerPreview) await waitFor(() => expect(worker.replies.has(request.id)).toBe(true));
      fullOpenId = request.id;
      layoutHeld = options.holdFullLayout === true;
      await act(async () => {
        if (workerPreview) worker.reply(request);
        else worker.workers.at(-1)!.release();
      });
      await waitFor(() => expect(harness.core.previewing).toBe(false));
      expect(harness.core.opening).toBe(true);
      expect(harness.core.replicaReady).toBe(false);
      return harness.core.session!;
    };
    const presentFull = async (full: YrsSession, startPeer = true) => {
      if (options.holdFullLayout) {
        await worker.received('completeLayout', fullOpenId);
      } else {
        act(() => harness.pipeline.runLayoutPipeline());
      }
      await waitFor(() => expect(harness.renderer.presentedEngine).toBe(full));
      act(() => harness.presentFrame());
      expect(harness.core.opening).toBe(false);
      if (!startPeer) return;
      await frames.settleAndIdle(harness.renderer.settledDisplayList(null, null, 'window'));
      await waitFor(() => expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(true));
    };
    const loadPeer = async (full: YrsSession) => {
      await act(async () => {
        worker.workers.at(-1)!.release();
        await awaitWorkerOpenReplica(full);
        await editor.current!.flushPendingInput();
      });
    };
    return {
      ...worker, frames, editor, view, preview, click, type, switchToFull, presentFull, loadPeer,
      mouseDown, mouseMove, mouseUp, pointAt, selections,
      reply(request: ResidentEngineWorkerRequest) {
        worker.reply(request);
        if (!options.holdFullLayout || request.type !== 'completeLayout') return;
        layoutHeld = false;
        for (const pending of worker.posted) {
          if (pending.type === 'sync' && pending.id > fullOpenId && worker.replies.has(pending.id)) {
            worker.reply(pending);
          }
        }
      },
      releaseInputQueries() {
        queriesReleased = true;
        releaseQueries();
      },
      get harness() { return harness; },
      get pendingQueries() { return previousInputQueries; },
      setReadOnly(readOnly: boolean, viewerSession = viewer) {
        view.rerender(<Editable {...props} readOnly={readOnly} viewer={viewerSession} />);
      },
      async replace() {
        if (options.holdFullLayout) layoutHeld = false;
        view.rerender(<Editable source={props.source.slice()} generation={2} />);
        const replacement = await waitForPreview(2);
        await waitFor(() => expect(worker.workers).toHaveLength(2));
        return replacement;
      },
      close() {
        view.unmount();
        releaseQueries();
        frames.restore();
      },
    };
  } catch (error) {
    view.unmount();
    releaseQueries();
    frames.restore();
    throw error;
  }
}

function pluginLoadGate() {
  let release!: () => void;
  const promise = new Promise<void>((resolve) => { release = resolve; });
  return { promise, release };
}

function workerLoadRecorder(load?: () => Promise<void>) {
  const events: DocxPluginEvent[] = [];
  const contexts: DocxPluginContext<null>[] = [];
  const plugin = defineDocxPlugin<null>({
    id: 'test.worker-load',
    createState: () => null,
    onEvent(context, event) {
      events.push(event);
      if (event.type !== 'load') return;
      contexts.push(context);
      return load?.();
    },
    commands: [{
      id: 'ready', label: 'Ready', mutatesDocument: false,
      execute: () => ({ ok: true, status: 'executed' }),
    }],
  });
  return {
    plugin, contexts,
    loads: () => events.filter((event) => event.type === 'load'),
    changes: () => events.filter((event) => event.type === 'document-change'),
  };
}

function openingPluginEditor(plugin: DocxPlugin, replay?: Promise<void>) {
  const restores: Array<() => void> = [];
  let view: ReturnType<typeof render> | undefined;
  const close = () => {
    try {
      view?.unmount();
    } finally {
      for (const restore of [...restores].reverse()) restore();
    }
  };
  try {
    const previousWorker = globalThis.Worker;
    const fonts = Object.getOwnPropertyDescriptor(document, 'fonts');
    restores.push(registerRestore(() => { globalThis.Worker = previousWorker; }));
    restores.push(registerRestore(() => {
      if (fonts) Object.defineProperty(document, 'fonts', fonts);
      else Reflect.deleteProperty(document, 'fonts');
    }));
    const frames = holdFrames(true);
    restores.push(frames.restore);
    let editorRef: React.RefObject<PagedEditorRef | null> | undefined;
    const bindPlugins = pluginHosts.useDocxPluginHost;
    const captureEditor = spyOn(pluginHosts, 'useDocxPluginHost').mockImplementation((options) => {
      editorRef = options.pagedEditorRef;
      return bindPlugins(options);
    });
    restores.push(registerRestore(() => captureEditor.mockRestore()));
    const peerReady = replicaHelpers.awaitWorkerOpenReplica;
    if (replay) {
      const wait = spyOn(replicaHelpers, 'awaitWorkerOpenReplica').mockImplementation((session) => {
        const peer = peerReady(session);
        return peer?.then(() => replay);
      });
      restores.push(registerRestore(() => wait.mockRestore()));
    }
    const fullOpen = (request: ResidentEngineWorkerRequest) =>
      request.type === 'open' && request.previewBlocks === undefined;
    const worker = installWorker({
      holdReply: (request) => fullOpen(request) || request.type === 'encodeState',
    });
    if (!document.fonts) Object.defineProperty(document, 'fonts', {
      value: {
        addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve(),
      },
      configurable: true,
    });
    const ref = createRef<DocxEditorRef>();
    const errors: unknown[] = [];
    let painted = false;
    const until = async (done: () => boolean) => {
      await waitFor(() => {
        act(() => { frames.run(); frames.runIdle(); });
        if (errors.length) throw new AggregateError(errors, 'Worker-open editor failed');
        expect(done()).toBe(true);
      }, { timeout: 5_000 });
    };
    const settle = async <T,>(promise: Promise<T>): Promise<T> => {
      let settled = false;
      void promise.then(() => { settled = true; }, () => { settled = true; });
      await until(() => settled);
      return promise;
    };
    const editor = () => editorRef?.current ?? null;
    const session = () => editor()?.getYrsSession() ?? null;
    const ready = () => !!ref.current?.commands.getDescriptor('plugin:test.worker-load/ready');
    const drain = async () => {
      for (let turn = 0; turn < 4; turn += 1) {
        await act(async () => {
          frames.run();
          frames.runIdle();
          await new Promise<void>((resolve) => setImmediate(resolve));
        });
      }
    };
    view = render(<DocxEditor ref={ref}
      documentBuffer={longBytes.slice().buffer as ArrayBuffer}
      mode="editing" experimentalWorkerOpen previewFirstPage plugins={[plugin]}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
      onFirstPagePainted={() => { painted = true; }} onError={(error) => errors.push(error)}
      onPluginError={(failure) => errors.push(failure.error)} />);
    return {
      ...worker, ref, view, until, settle, ready, drain, editor, session, close,
      async preview() {
        await until(() => painted && session()?.isDisplayOnly() === true);
        await until(() => worker.posted.some((request) => fullOpen(request) && worker.replies.has(request.id)));
        return session()!;
      },
      type(key: string) {
        const textarea = view!.getByTestId('yrs-input');
        fireEvent.keyDown(textarea, { key });
        fireEvent.input(textarea, { target: { value: key } });
      },
      async open() {
        const request = worker.posted.find(fullOpen)!;
        act(() => worker.reply(request));
        await until(() => session() !== null && session()?.isDisplayOnly() === false);
        await until(() => worker.posted.some((request) =>
          request.type === 'encodeState' && worker.replies.has(request.id)));
        return session()!;
      },
      async hydrate(full: YrsSession) {
        const request = worker.posted.find((request) => request.type === 'encodeState')!;
        act(() => worker.reply(request));
        const peer = peerReady(full);
        if (!peer) throw new Error('The editor peer was not scheduled');
        await settle(peer);
        await until(() => !replicaHelpers.workerOpenReplicaPending(full));
      },
    };
  } catch (error) {
    close();
    throw error;
  }
}

async function boundedReplay<T>(promise: Promise<T>): Promise<T> {
  let timer!: ReturnType<typeof setTimeout>;
  try {
    return await Promise.race([
      promise,
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new Error('Worker replay did not settle')), 5_000);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

async function openingReplayEditor() {
  const previousWorker = globalThis.Worker;
  const fonts = Object.getOwnPropertyDescriptor(document, 'fonts');
  const restore = registerRestore(() => {
    globalThis.Worker = previousWorker;
    if (fonts) Object.defineProperty(document, 'fonts', fonts);
    else Reflect.deleteProperty(document, 'fonts');
  });
  try {
    const opened = await openingEditor(true, false, { residentInput: true, publishLayout: true });
    return {
      ...opened,
      get harness() { return opened.harness; },
      close() {
        try {
          opened.close();
        } finally {
          opened.frames.restore();
          restore();
        }
      },
    };
  } catch (error) {
    restore();
    throw error;
  }
}

async function settledReplayLayout(opened: Awaited<ReturnType<typeof openingReplayEditor>>, session: YrsSession) {
  await waitFor(() => {
    act(() => opened.frames.run());
    expect(opened.harness.errors).toEqual([]);
    expect(opened.harness.renderer.queries?.isReady()).toBe(true);
    expect(sourceVersionOf(opened.harness.renderer.queries)).toBe(session.version());
  }, { timeout: 5_000 });
  for (let turn = 0; turn < 4; turn += 1) {
    await act(async () => {
      opened.frames.run();
      await new Promise<void>((resolve) => setImmediate(resolve));
    });
  }
}

function holdIdle() {
  const request = globalThis.requestIdleCallback;
  const cancel = globalThis.cancelIdleCallback;
  const callbacks = new Map<number, { callback: IdleRequestCallback; options?: IdleRequestOptions }>();
  let nextId = 0;
  globalThis.requestIdleCallback = (callback, options) => {
    callbacks.set(++nextId, { callback, options });
    return nextId;
  };
  globalThis.cancelIdleCallback = (id) => { callbacks.delete(id); };
  return {
    callbacks,
    run() {
      const pending = [...callbacks.values()];
      callbacks.clear();
      for (const { callback } of pending) callback({ didTimeout: false, timeRemaining: () => 50 });
    },
    restore: registerRestore(() => {
      globalThis.requestIdleCallback = request;
      globalThis.cancelIdleCallback = cancel;
    }),
  };
}

function holdHydrationTasks() {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'scheduler');
  const tasks: Array<() => void> = [];
  Object.defineProperty(globalThis, 'scheduler', {
    configurable: true,
    value: { yield: () => new Promise<void>((resolve) => tasks.push(resolve)) },
  });
  return {
    tasks,
    async run() {
      const task = tasks.shift();
      if (!task) throw new Error('No pending hydration task');
      task();
      await Promise.resolve();
    },
    restore: registerRestore(() => {
      if (descriptor) Object.defineProperty(globalThis, 'scheduler', descriptor);
      else Reflect.deleteProperty(globalThis, 'scheduler');
    }),
  };
}

function holdFontEditorReadiness(trace: string[]) {
  const progress = replicaHelpers.notifyWorkerOpenLayoutProgress;
  const hold = spyOn(replicaHelpers, 'notifyWorkerOpenLayoutProgress').mockImplementation((session, kind) => {
    progress(session, !session.isDisplayOnly() && kind === 'complete' ? 'provisional' : kind);
  });
  const createHost = pluginHostFactories.createDocxPluginHost;
  let readLayout: Parameters<typeof createHost>[0]['layout'] = () => ({
    queries: null, complete: false, failed: false,
  });
  const capture = spyOn(pluginHostFactories, 'createDocxPluginHost').mockImplementation((access) => {
    readLayout = access.layout;
    const host = createHost(access);
    const open = host.open;
    host.open = (session) => {
      trace.push('plugin-open');
      open(session);
    };
    return host;
  });
  return {
    complete: (session: YrsSession) => progress(session, 'complete'),
    layout: () => readLayout(),
    restore: registerRestore(() => {
      capture.mockRestore();
      hold.mockRestore();
    }),
  };
}

async function pendingFontEditor() {
  const trace: string[] = [];
  const readiness = holdFontEditorReadiness(trace);
  const recorder = workerLoadRecorder();
  const opened = openingPluginEditor(recorder.plugin);
  const fallback = holdLayoutFallback();
  let tasks: ReturnType<typeof holdHydrationTasks> | undefined;
  let restoreRequirements = () => {};
  const close = () => {
    try {
      opened.close();
    } finally {
      restoreRequirements();
      tasks?.restore();
      fallback.restore();
      readiness.restore();
    }
  };
  try {
    await opened.preview();
    const full = await opened.open();
    await opened.drain();
    const readRequirements = full.layoutFontRequirementsJson;
    const requirements = spyOn(full, 'layoutFontRequirementsJson').mockImplementation((input) => {
      trace.push('font-refresh');
      return readRequirements(input);
    });
    restoreRequirements = registerRestore(() => requirements.mockRestore());
    const initializedPosts = opened.posted.length;
    tasks = holdHydrationTasks();
    const state = opened.posted.find((request) => request.type === 'encodeState')!;
    act(() => opened.reply(state));
    await opened.until(() => tasks!.tasks.length === 1);
    expect(replicaHelpers.workerOpenReplicaStarted(full)).toBe(true);
    expect(replicaHelpers.workerOpenReplicaPending(full)).toBe(true);
    expect(replicaHelpers.workerOpenReplicaHydrating(full)).toBe(true);
    return { ...opened, full, trace, readiness, requirements, initializedPosts, tasks, fallback, close };
  } catch (error) {
    close();
    throw error;
  }
}

function holdPeerFallback(options: { allTimers?: boolean } = {}) {
  const schedule = globalThis.setTimeout;
  const unschedule = globalThis.clearTimeout;
  const timers = new Map<number, { at: number; callback: () => void }>();
  let now = 0;
  let nextId = 0;
  globalThis.setTimeout = ((...input: Parameters<typeof setTimeout>) => {
    const [callback, delay, ...args] = input;
    if ((options.allTimers || delay === 10_000) && typeof callback === 'function') {
      const id = --nextId;
      timers.set(id, { at: now + (delay ?? 0), callback: () => callback(...args) });
      return id as unknown as ReturnType<typeof setTimeout>;
    }
    return schedule(callback, delay, ...args);
  }) as typeof setTimeout;
  globalThis.clearTimeout = ((id?: ReturnType<typeof setTimeout>) => {
    if (!timers.delete(id as unknown as number)) unschedule(id);
  }) as typeof clearTimeout;
  return {
    timers,
    get now() { return now; },
    advance(ms: number) {
      if (!options.allTimers) {
        now += ms;
        for (const [id, timer] of [...timers]) {
          if (timer.at > now) continue;
          timers.delete(id);
          timer.callback();
        }
        return;
      }
      const target = now + ms;
      for (;;) {
        const next = [...timers].filter(([, timer]) => timer.at <= target)
          .sort((a, b) => a[1].at - b[1].at)[0];
        if (!next) break;
        const [id, timer] = next;
        now = timer.at;
        timers.delete(id);
        timer.callback();
      }
      now = target;
    },
    restore: registerRestore(() => {
      globalThis.setTimeout = schedule;
      globalThis.clearTimeout = unschedule;
    }),
  };
}

function holdLayoutFallback() {
  const schedule = globalThis.setTimeout;
  const unschedule = globalThis.clearTimeout;
  const timers = new Map<number, { at: number; callback: () => void }>();
  let now = 0;
  let nextId = 0;
  globalThis.setTimeout = ((...input: Parameters<typeof setTimeout>) => {
    const [callback, delay, ...args] = input;
    if ((delay === 3000 || delay === 30_000) && typeof callback === 'function') {
      const id = --nextId;
      timers.set(id, { at: now + delay, callback: () => callback(...args) });
      return id as unknown as ReturnType<typeof setTimeout>;
    }
    return schedule(callback, delay, ...args);
  }) as typeof setTimeout;
  globalThis.clearTimeout = ((id?: ReturnType<typeof setTimeout>) => {
    if (!timers.delete(id as unknown as number)) unschedule(id);
  }) as typeof clearTimeout;
  return {
    timers,
    advance(ms: number) {
      const target = now + ms;
      for (;;) {
        const next = [...timers].filter(([, timer]) => timer.at <= target)
          .sort((a, b) => a[1].at - b[1].at)[0];
        if (!next) break;
        const [id, timer] = next;
        now = timer.at;
        timers.delete(id);
        timer.callback();
      }
      now = target;
    },
    restore: registerRestore(() => {
      globalThis.setTimeout = schedule;
      globalThis.clearTimeout = unschedule;
    }),
  };
}

async function finishHeldHydration(
  opened: Awaited<ReturnType<typeof openingEditor>>,
  tasks: ReturnType<typeof holdHydrationTasks>
) {
  await waitFor(() => expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(true));
  await act(async () => {
    while (tasks.tasks.length) await tasks.run();
  });
  expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
  await act(async () => opened.workers.at(-1)!.release());
  await waitFor(() => expect(tasks.tasks).toHaveLength(1));
  await act(async () => tasks.run());
  await act(async () => tasks.run());
  tasks.restore();
}

function stubDocumentVisibility(initial: 'visible' | 'hidden') {
  const descriptor = Object.getOwnPropertyDescriptor(document, 'visibilityState');
  let state = initial;
  Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => state });
  return {
    hide() {
      state = 'hidden';
      document.dispatchEvent(new Event('visibilitychange'));
    },
    restore: registerRestore(() => {
      if (descriptor) Object.defineProperty(document, 'visibilityState', descriptor);
      else Reflect.deleteProperty(document, 'visibilityState');
    }),
  };
}

function expectStatePrefetchAfterPageBuilds(posted: readonly ResidentEngineWorkerRequest[]): void {
  const types = posted.map((request) => request.type);
  expect(types.filter((type) => type === 'encodeState')).toHaveLength(1);
  expect(types.indexOf('encodeState')).toBeGreaterThan(types.lastIndexOf('buildPages'));
}

async function openingViewer() {
  const worker = installWorker({ holdReply: (request) => request.type === 'bootstrap' });
  const frames = holdFrames(true);
  const fonts = Object.getOwnPropertyDescriptor(document, 'fonts');
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
    configurable: true,
  });
  const editor = createRef<PagedEditorRef>();
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function Viewer() {
    harness = useHarness({ ...initialProps, source: longBytes, viewer: true, readOnly: true });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        readOnly viewerDocumentRead={harness.renderer.readWorkerDocument}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        onLayoutComputed={harness.renderer.onLayoutComputed}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries} />
    </>;
  }
  const view = render(<Viewer />);
  const close = () => {
    try {
      view.unmount();
    } finally {
      frames.restore();
      if (fonts) Object.defineProperty(document, 'fonts', fonts);
      else Reflect.deleteProperty(document, 'fonts');
    }
  };
  try {
    await waitFor(() => expect(harness.host).not.toBeNull());
    const session = harness.core.session!;
    act(() => harness.pipeline.runLayoutPipeline());
    const bootstrap = await worker.received('bootstrap');
    return {
      ...worker, frames, editor, view, session, close,
      get harness() { return harness; },
      click() {
        const canvas = canvasHost.current!.firstElementChild!;
        fireEvent.mouseDown(canvas, { clientX: 0, clientY: 0, button: 0 });
        fireEvent.mouseUp(window, { clientX: 0, clientY: 0, button: 0 });
      },
      async present() {
        act(() => worker.reply(bootstrap));
        await waitFor(() => expect(harness.renderer.presentedEngine).toBe(session));
        act(() => harness.presentFrame());
        return session;
      },
    };
  } catch (error) {
    close();
    throw error;
  }
}

test('worker-open plugins retry load once when a held key replays during async load', async () => {
  const replay = pluginLoadGate();
  const load = pluginLoadGate();
  const recorder = workerLoadRecorder(() => load.promise);
  const opened = openingPluginEditor(recorder.plugin, replay.promise);
  try {
    const preview = await opened.preview();
    const previewVersion = preview.version();
    opened.type('A');
    expect(opened.editor()!.hasPendingInput()).toBe(true);
    expect(preview.version()).toBe(previewVersion);
    expect(recorder.loads()).toEqual([]);
    const full = await opened.open();
    expect(full.storyIds()).toEqual([]);
    await opened.hydrate(full);
    await opened.until(() => recorder.loads().length === 1);
    const initial = recorder.loads()[0]!;
    expect(initial.reason).toBe('loaded');
    expect(initial.version).toBe(full.version());
    expect(recorder.contexts[0]!.signal.aborted).toBe(false);
    expect(opened.ready()).toBe(false);
    act(() => replay.release());
    await opened.settle(opened.editor()!.flushPendingInput());
    const editedVersion = full.version();
    expect(editedVersion).not.toBe(initial.version);
    expect(full.paragraphs('body')[0]!.text).toBe('AFirst paragraph');
    expect(recorder.contexts[0]!.signal.aborted).toBe(true);
    expect(recorder.loads()).toHaveLength(1);
    act(() => load.release());
    await opened.until(opened.ready);
    await opened.drain();
    expect(recorder.loads()).toEqual([
      initial,
      { type: 'load', generation: initial.generation, reason: 'loaded', version: editedVersion },
    ]);
    expect(recorder.contexts.map((context) => context.snapshot.version))
      .toEqual([initial.version, editedVersion]);
    expect(recorder.changes()).toEqual([]);
  } finally {
    replay.release();
    load.release();
    opened.close();
  }
}, 30_000);

test('worker-open typing after plugin load settles emits one document-change and no further load', async () => {
  const recorder = workerLoadRecorder();
  const opened = openingPluginEditor(recorder.plugin);
  try {
    await opened.preview();
    const full = await opened.open();
    await opened.hydrate(full);
    await opened.until(opened.ready);
    expect(recorder.loads()).toHaveLength(1);
    expect(recorder.changes()).toEqual([]);
    const initial = recorder.loads()[0]!;
    opened.type('B');
    await opened.settle(opened.editor()!.flushPendingInput());
    await opened.until(() => recorder.changes().length === 1);
    await opened.drain();
    expect(full.paragraphs('body')[0]!.text).toBe('BFirst paragraph');
    expect(full.version()).not.toBe(initial.version);
    expect(recorder.loads()).toEqual([initial]);
    expect(recorder.changes()).toEqual([
      { type: 'document-change', generation: initial.generation, version: full.version() },
    ]);
  } finally {
    opened.close();
  }
}, 30_000);

test('an eager worker-open editor peer becoming ready delivers only the initial plugin load', async () => {
  const recorder = workerLoadRecorder();
  const opened = openingPluginEditor(recorder.plugin);
  try {
    await opened.preview();
    const full = await opened.open();
    expect(replicaHelpers.workerOpenReplicaPending(full)).toBe(true);
    expect(full.storyIds()).toEqual([]);
    await opened.hydrate(full);
    await opened.until(opened.ready);
    await opened.drain();
    expect(replicaHelpers.workerOpenReplicaPending(full)).toBe(false);
    expect(full.paragraphs('body')[0]!.text).toBe('First paragraph');
    expect(recorder.loads()).toEqual([
      { type: 'load', generation: recorder.loads()[0]!.generation, reason: 'loaded', version: full.version() },
    ]);
    expect(recorder.contexts).toHaveLength(1);
    expect(recorder.changes()).toEqual([]);
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
  } finally {
    opened.close();
  }
}, 30_000);

test.each(['held keys', 'held composition'] as const)(
  'flush during opening waits for %s to land in the full session', async (kind) => {
    const opened = await openingEditor(true, false, { publishLayout: true });
    try {
      const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
      opened.click(6);
      if (kind === 'held keys') fireEvent.input(textarea, { target: { value: 'A' } });
      else {
        fireEvent.compositionStart(textarea);
        textarea.value = '日本';
      }
      let done = false;
      let seen: string[] = [];
      const flush = opened.editor.current!.flushPendingInput().then(() => {
        seen = opened.editor.current!.getYrsSession()!.paragraphs('body').map((paragraph) => paragraph.text);
        done = true;
      });
      await act(async () => { await Promise.resolve(); });
      expect(done).toBe(false);
      const full = await opened.switchToFull();
      expect(done).toBe(false);
      if (kind === 'held keys') fireEvent.input(textarea, { target: { value: 'B' } });
      await opened.presentFull(full);
      if (kind === 'held composition') fireEvent.compositionEnd(textarea, { data: '日本' });
      await act(async () => { await Promise.resolve(); });
      expect(done).toBe(false);
      expect(full.storyIds()).toEqual([]);
      await opened.loadPeer(full);
      await act(async () => { await flush; });
      expect(done).toBe(true);
      expect(seen[0]).toBe(kind === 'held keys' ? 'FirstAB paragraph' : 'First日本 paragraph');
      expect(full.selection()?.head.offset).toBe(7);
      expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
      expect(opened.harness.mainOpens).toEqual([false]);
    } finally {
      opened.close();
    }
  }
);

test('flush during opening rejects when another document replaces held input', async () => {
  const opened = await openingEditor(false);
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    fireEvent.input(textarea, { target: { value: 'discard' } });
    const flush = opened.editor.current!.flushPendingInput().catch((error) => error);
    await opened.replace();
    expect(await flush).toEqual(new Error('The document changed while flushing input'));
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
  } finally {
    opened.close();
  }
});

test('flush with nothing held during opening keeps preview and replica readiness behavior', async () => {
  const opened = await openingEditor();
  try {
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    await act(async () => { await opened.editor.current!.flushPendingInput(); });
    expect(opened.harness.core.previewing).toBe(true);
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    const full = await opened.switchToFull();
    let done = false;
    const flush = opened.editor.current!.flushPendingInput().then(() => { done = true; });
    await act(async () => { await Promise.resolve(); });
    expect(done).toBe(false);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    await opened.presentFull(full);
    expect(done).toBe(false);
    await opened.loadPeer(full);
    await act(async () => { await flush; });
    expect(done).toBe(true);
    expect(full.paragraphs('body')[0].text).toBe('First paragraph');
  } finally {
    opened.close();
  }
});

test('keys typed during opening land in order, none dropped', async () => {
  const opened = await openingEditor();
  const insert = spyOn(opened.preview, 'insertText');
  const split = spyOn(opened.preview, 'splitParagraph');
  const remove = spyOn(opened.preview, 'deleteRange');
  const select = spyOn(opened.preview, 'setSelection');
  try {
    const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
    const version = opened.preview.version();
    expect(opened.preview.storyIds()).toEqual([]);
    expect(textarea.readOnly).toBe(false);
    opened.click(6);
    expect(document.activeElement).toBe(textarea);
    fireEvent.input(textarea, { target: { value: 'A' } });
    fireEvent.keyDown(textarea, { key: 'Enter' });
    fireEvent.input(textarea, { target: { value: 'BC' } });
    fireEvent.keyDown(textarea, { key: 'Backspace' });
    let clipboard = 'P\r\nQ';
    fireEvent.paste(textarea, { clipboardData: { getData: () => clipboard } });
    clipboard = 'invalidated';
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(opened.preview.version()).toBe(version);
    expect(insert).not.toHaveBeenCalled();
    expect(split).not.toHaveBeenCalled();
    expect(remove).not.toHaveBeenCalled();
    expect(select).not.toHaveBeenCalled();
    const full = await opened.switchToFull();
    const load = spyOn(full, 'loadState');
    try {
      expect(full.storyIds()).toEqual([]);
      await opened.presentFull(full);
      fireEvent.input(textarea, { target: { value: 'R' } });
      expect(load).not.toHaveBeenCalled();
      await opened.loadPeer(full);
      expect(full.paragraphs('body').slice(0, 3).map((paragraph) => paragraph.text))
        .toEqual(['FirstA', 'BP', 'QR paragraph']);
      expect(full.selection()?.head).toEqual({
        story: 'body', paraId: full.paragraphs('body')[2].paraId, offset: 2,
      });
      act(() => opened.harness.presentFrame());
      act(() => opened.frames.run());
      act(() => opened.frames.run());
      act(() => opened.frames.runIdle());
      await act(async () => { await opened.editor.current!.flushPendingInput(); });
      expect(full.paragraphs('body').slice(0, 3).map((paragraph) => paragraph.text))
        .toEqual(['FirstA', 'BP', 'QR paragraph']);
      expect(load).toHaveBeenCalledTimes(1);
      expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
      expect(opened.harness.mainOpens).toEqual([false]);
      expect(insert).not.toHaveBeenCalled();
      expect(split).not.toHaveBeenCalled();
      expect(remove).not.toHaveBeenCalled();
      expect(select).not.toHaveBeenCalled();
      expect(textarea.value).toBe('');
    } finally {
      load.mockRestore();
    }
  } finally {
    insert.mockRestore();
    split.mockRestore();
    remove.mockRestore();
    select.mockRestore();
    opened.close();
  }
});

test('a key typed right after the first page paints is kept', async () => {
  const opened = await openingEditor(true, false, {
    delayQueries: true,
    onFirstPagePainted: ({ click, type }) => {
      click(6);
      type('QZXJ');
    },
  });
  const insertPreview = spyOn(opened.preview, 'insertText');
  try {
    const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
    expect(textarea.readOnly).toBe(false);
    expect(document.activeElement).toBe(textarea);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(opened.preview.storyIds()).toEqual([]);
    const full = await opened.switchToFull();
    const load = spyOn(full, 'loadState');
    try {
      expect(opened.view.getByTestId('yrs-input')).toBe(textarea);
      expect(document.activeElement).toBe(textarea);
      await opened.presentFull(full);
      expect(load).not.toHaveBeenCalled();
      opened.releaseInputQueries();
      await opened.loadPeer(full);
      expect(full.paragraphs('body')[0].text).toBe('FirstQZXJ paragraph');
      expect(full.selection()?.head.offset).toBe(9);
      act(() => opened.harness.presentFrame());
      act(() => opened.frames.run());
      act(() => opened.frames.run());
      act(() => opened.frames.runIdle());
      await act(async () => { await opened.editor.current!.flushPendingInput(); });
      expect(full.paragraphs('body')[0].text).toBe('FirstQZXJ paragraph');
      expect(load).toHaveBeenCalledTimes(1);
      expect(insertPreview).not.toHaveBeenCalled();
      expect(opened.harness.mainOpens).toEqual([false]);
      expect(opened.harness.errors).toEqual([]);
    } finally {
      load.mockRestore();
    }
  } finally {
    insertPreview.mockRestore();
    opened.close();
  }
});

test('two bursts typed milliseconds apart around first paint both land, in order', async () => {
  const opened = await openingEditor(true, false, {
    delayQueries: true,
    onFirstPagePainted: ({ click, type }) => {
      click(6);
      type('QZXJ');
      click(6);
      type('JWKV');
    },
  });
  try {
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    const full = await opened.switchToFull();
    const load = spyOn(full, 'loadState');
    try {
      await opened.presentFull(full);
      opened.releaseInputQueries();
      await opened.loadPeer(full);
      expect(full.paragraphs('body')[0].text).toBe('FirstQZXJJWKV paragraph');
      expect(full.selection()?.head.offset).toBe(13);
      await act(async () => { await opened.editor.current!.flushPendingInput(); });
      expect(full.paragraphs('body')[0].text).toBe('FirstQZXJJWKV paragraph');
      expect(load).toHaveBeenCalledTimes(1);
      expect(opened.harness.errors).toEqual([]);
    } finally {
      load.mockRestore();
    }
  } finally {
    opened.close();
  }
});

test('held text runs separated by a click replay with one worker layout request and match ready typing', async () => {
  let replayedText!: string[];
  let replayedOffset!: number;
  const held = await openingReplayEditor();
  try {
    held.click(6);
    held.type('AB');
    held.click(2);
    held.type('CD');
    const full = await boundedReplay(held.switchToFull());
    await boundedReplay(held.presentFull(full));
    const start = held.posted.length;
    await boundedReplay(held.loadPeer(full));
    await settledReplayLayout(held, full);
    const requests = held.posted.slice(start);
    expect(requests.filter((request) =>
      request.type === 'sync' || request.type === 'bootstrap' ||
      request.type === 'applyInput' || request.type === 'applyDelete'
    ).map((request) => request.type)).toEqual(['sync']);
    expect(requests.filter((request) => request.type === 'applyUpdate')).toHaveLength(4);
    expect(full.paragraphs('body')[0]!.text).toBe('FCDirstAB paragraph');
    expect(held.editor.current!.hasPendingInput()).toBe(false);
    replayedText = full.paragraphs('body').map((paragraph) => paragraph.text);
    replayedOffset = full.selection()!.head.offset;
    held.type('E');
    await boundedReplay(held.editor.current!.flushPendingInput());
    expect(held.posted.slice(start).filter((request) => request.type === 'applyInput'))
      .toMatchObject([{ text: 'E' }]);
  } finally {
    held.close();
  }
  const ready = await openingReplayEditor();
  try {
    const full = await boundedReplay(ready.switchToFull());
    await boundedReplay(ready.presentFull(full));
    await boundedReplay(ready.loadPeer(full));
    ready.click(6);
    ready.type('AB');
    await boundedReplay(ready.editor.current!.flushPendingInput());
    await settledReplayLayout(ready, full);
    ready.click(2);
    ready.type('CD');
    await boundedReplay(ready.editor.current!.flushPendingInput());
    expect(full.paragraphs('body').map((paragraph) => paragraph.text)).toEqual(replayedText);
    expect(full.selection()!.head.offset).toBe(replayedOffset);
    expect(ready.harness.errors).toEqual([]);
  } finally {
    ready.close();
  }
}, 30_000);

test('clicks interleaved with held and hydrating input keep their original positions', async () => {
  const opened = await openingEditor();
  try {
    opened.click(6);
    opened.type('QZXJ');
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    opened.click(2);
    opened.type('JWKV');
    await opened.loadPeer(full);
    expect(full.paragraphs('body')[0].text).toBe('FJWKVirstQZXJ paragraph');
    expect(full.selection()?.head.offset).toBe(5);
    expect(opened.harness.errors).toEqual([]);
    expect(opened.harness.mainOpens).toEqual([false]);
  } finally {
    opened.close();
  }
});

test('a click during hydration preserves A, ArrowLeft, B in order and places the final caret', async () => {
  const opened = await openingEditor();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    expect(opened.harness.core.replicaReady).toBe(false);
    opened.click(6);
    opened.type('A');
    fireEvent.keyDown(opened.view.getByTestId('yrs-input'), { key: 'ArrowLeft' });
    opened.type('B');
    opened.click(2);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    await opened.loadPeer(full);
    expect(full.paragraphs('body')[0].text).toBe('FirstBA paragraph');
    const caret = { story: 'body', paraId: full.paragraphs('body')[0].paraId, offset: 1 };
    expect(full.selection()).toEqual({ anchor: caret, head: caret });
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test('table Tab during hydration follows held input and moves to the next cell', async () => {
  const zip = await JSZip.loadAsync(longBytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  const cells = ['Left', 'Right'].map((text) =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
  ).join('');
  zip.file('word/document.xml', xml.replace(
    '<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>',
    `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`
  ));
  const source = await zip.generateAsync({ type: 'uint8array' });
  const direct = await createYrsSession();
  sessions.push(direct);
  direct.openDocx(source, true);
  const projection = createYrsPositionProjection(direct, 'body')!;
  const left = 'body:t0:r0c0';
  const right = 'body:t0:r0c1';
  const position = projection.positionForLoc({ story: left, paraId: direct.paragraphs(left)[0]!.paraId, offset: 1 })!;
  const opened = await openingEditor(true, false, { source });
  try {
    opened.click(position);
    opened.type('A');
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    expect(opened.harness.core.opening).toBe(false);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(full.storyIds()).toEqual([]);
    const textarea = opened.view.getByTestId('yrs-input');
    expect(fireEvent.keyDown(textarea, { key: 'Tab' })).toBe(false);
    opened.type('B');
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(full.storyIds()).toEqual([]);
    await opened.loadPeer(full);
    expect(full.paragraphs(left)[0]!.text).toBe('LAeft');
    expect(full.paragraphs(right)[0]!.text).toBe('BRight');
    expect(full.selection()?.head.story).toBe(right);
    expect(full.selection()?.head.offset).toBe(1);
    expect(full.cellSelection()?.head.column).toBe(1);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test('a ready click binds its position before an outstanding resident edit reply', async () => {
  const zip = await JSZip.loadAsync(await longFixture(1));
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('First paragraph', 'abcdef'));
  const source = await zip.generateAsync({ type: 'uint8array' });
  const opened = await openingEditor(true, false, {
    source, residentInput: true,
    holdReply: (request) => request.type === 'applyInput' && request.text === 'X',
  });
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    await opened.loadPeer(full);
    expect(full.paragraphs('body')[0].text).toBe('abcdef');
    opened.click(1);
    opened.type('X');
    const request = await opened.received('applyInput');
    expect(request).toMatchObject({ text: 'X', selection: { head: { offset: 0 } } });
    expect(opened.replies.has(request.id)).toBe(true);
    expect(opened.responses.get(request)).toMatchObject({ ok: true, frame: expect.any(ArrayBuffer) });
    expect(full.paragraphs('body')[0].text).toBe('abcdef');
    expect(opened.harness.renderer.queries!.isReady()).toBe(true);
    opened.click(4);
    opened.type('Y');
    expect(full.selection()?.head.offset).toBe(0);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    await act(async () => {
      opened.reply(request);
      await opened.editor.current!.flushPendingInput();
    });
    expect(full.paragraphs('body')[0].text).toBe('XabcYdef');
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test('a click awaiting replacement geometry keeps its paragraph after the canvas moves', async () => {
  const opened = await openingEditor(true, false, { delayQueries: true });
  try {
    const queries = opened.pendingQueries!;
    const point = opened.pointAt(6);
    expect(queries.isReady()).toBe(false);
    expect(queries.hitTestRegions(0, point.clientX, point.clientY)).toBeNull();
    opened.click(6);
    opened.type('Q');
    const nextCaret = opened.harness.renderer.queries!.caretRect(18)!;
    const size = queries.pageSize(0)!;
    const top = point.clientY - nextCaret.y - nextCaret.height / 2;
    const canvas = opened.view.container.querySelector('.canvas-page')!;
    canvas.getBoundingClientRect = () => ({ left: 0, top, right: size.width,
      bottom: top + size.height, ...size }) as DOMRect;
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    expect(opened.pendingQueries).not.toBe(queries);
    expect(queries.isReady()).toBe(false);
    opened.releaseInputQueries();
    await queries.whenReady();
    expect(queries.isReady()).toBe(true);
    expect(queries.hitTestRegions(0, point.clientX, point.clientY)).toBeNull();
    await opened.pendingQueries!.whenReady();
    expect(opened.pendingQueries!.hitTestRegions(0, point.clientX, point.clientY - top)?.pos).toBeGreaterThan(16);
    await opened.loadPeer(full);
    expect(full.paragraphs('body')[0].text).toBe('FirstQ paragraph');
    expect(full.paragraphs('body')[1].text).toBe('Paragraph 1');
    expect(full.selection()?.head.offset).toBe(6);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test('a queued drag across cells in one table publishes the cell selection', async () => {
  const zip = await JSZip.loadAsync(longBytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  const cells = ['Left', 'Right'].map((text) =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
  ).join('');
  zip.file('word/document.xml', xml.replace(
    '<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>',
    `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`
  ));
  const source = await zip.generateAsync({ type: 'uint8array' });
  const direct = await createYrsSession();
  sessions.push(direct);
  direct.openDocx(source, true);
  const projection = createYrsPositionProjection(direct, 'body')!;
  const position = (story: string) => projection.positionForLoc({
    story, paraId: direct.paragraphs(story)[0].paraId, offset: 1,
  })!;
  const anchorPosition = position('body:t0:r0c0');
  const headPosition = position('body:t0:r0c1');
  const anchorTarget = projection.targetAt(anchorPosition);
  const headTarget = projection.targetAt(headPosition);
  expect(anchorTarget.story).not.toBe(headTarget.story);
  const range = { anchor: anchorTarget.cell!, head: headTarget.cell! };
  const opened = await openingEditor(true, false, { source, delayQueries: true });
  try {
    opened.mouseDown(anchorPosition);
    opened.mouseMove(headPosition);
    opened.mouseUp(headPosition);
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    opened.releaseInputQueries();
    await opened.loadPeer(full);
    expect(full.cellSelection()).toEqual(range);
    expect(full.selection()?.head.story).toBe(anchorTarget.story);
    expect(opened.selections).toContainEqual(range);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test.each(['held keys', 'held composition'] as const)(
  'becoming read-only during opening discards %s and rejects its flush', async (kind) => {
    const opened = await openingEditor(true, false, { delayQueries: true });
    try {
      const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
      opened.click(6);
      if (kind === 'held keys') opened.type('discard');
      else {
        fireEvent.compositionStart(textarea);
        textarea.value = '日本';
      }
      const flush = opened.editor.current!.flushPendingInput().catch((error) => error);
      expect(opened.editor.current!.hasPendingInput()).toBe(true);
      opened.setReadOnly(true);
      expect(opened.harness.core.session).toBe(opened.preview);
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
      expect(await flush).toEqual(new Error('The document changed while flushing input'));
      if (kind === 'held composition') fireEvent.compositionEnd(textarea, { data: '日本' });
      opened.setReadOnly(false);
      const full = await opened.switchToFull();
      await opened.presentFull(full);
      opened.releaseInputQueries();
      await opened.loadPeer(full);
      expect(full.paragraphs('body')[0].text).toBe('First paragraph');
      opened.type('C');
      await act(async () => { await opened.editor.current!.flushPendingInput(); });
      expect(full.paragraphs('body')[0].text).toBe('CFirst paragraph');
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
      expect(opened.harness.errors).toEqual([]);
    } finally {
      opened.close();
    }
  }
);

test('switching to viewing before eager hydration settles transferred input and its flush', async () => {
  const opened = await openingEditor();
  try {
    opened.click(1);
    opened.type('Q');
    const full = await opened.switchToFull();
    const load = spyOn(full, 'loadState');
    registerRestore(() => load.mockRestore());
    await opened.presentFull(full, false);
    await opened.frames.settleAndIdle(opened.harness.renderer.settledDisplayList(null, null, 'window'));
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.harness.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    let flushed = false;
    const flush = opened.editor.current!.flushPendingInput().then(() => { flushed = true; });
    opened.setReadOnly(true, true);
    expect(opened.editor.current!.isWorkerViewer()).toBe(false);
    expect((opened.view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(flushed).toBe(false);
    await opened.frames.settleAndIdle(opened.harness.renderer.settledDisplayList(null, null, 'window'));
    await waitFor(() => expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => {
      opened.workers.at(-1)!.release();
      await flush;
    });
    expect(flushed).toBe(true);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(full.paragraphs('body')[0].text).toBe('QFirst paragraph');
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test('a queued drag extends from its resolved anchor after an earlier insertion', async () => {
  const opened = await openingEditor(true, false, { delayQueries: true });
  try {
    opened.click(1);
    opened.type('Q');
    opened.mouseDown(6);
    opened.mouseMove(9);
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    opened.releaseInputQueries();
    await opened.loadPeer(full);
    expect(full.paragraphs('body')[0].text).toBe('QFirst paragraph');
    expect(full.selection()?.anchor.offset).toBe(6);
    expect(full.selection()?.head.offset).toBe(9);
    opened.mouseMove(12);
    act(() => opened.frames.run());
    opened.mouseUp(12);
    await act(async () => { await opened.editor.current!.flushPendingInput(); });
    expect(full.selection()?.anchor.offset).toBe(6);
    expect(full.selection()?.head.offset).toBe(11);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test('a ready drag queued behind navigation replays its final range once', async () => {
  const { opened, pointAt, startNavigation, releaseNavigation } = await readyEditorWithPendingNavigation();
  const select = spyOn(opened.session, 'setSelection');
  try {
    await startNavigation();
    const before = opened.session.selection();
    fireEvent.mouseDown(opened.canvas, pointAt(6));
    fireEvent.mouseMove(window, pointAt(10));
    fireEvent.mouseUp(window, pointAt(10));
    expect(opened.session.selection()).toEqual(before);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    await act(async () => {
      releaseNavigation();
      await opened.editor.current!.flushPendingInput();
    });
    const anchor = { story: 'body', paraId: opened.session.paragraphs('body')[0].paraId, offset: 5 };
    const head = { ...anchor, offset: 9 };
    const assertReplay = () => {
      expect(opened.session.selection()).toEqual({ anchor, head });
      expect(select.mock.calls.filter(([start, end]) => start.offset === 5 && end?.offset === 9)).toHaveLength(1);
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
    };
    assertReplay();
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
      await opened.editor.current!.flushPendingInput();
    });
    assertReplay();
    expect(opened.harness.errors).toEqual([]);
  } finally {
    releaseNavigation();
    select.mockRestore();
    opened.close();
  }
});

test('a ready image click queued behind navigation selects the image instead of a drag', async () => {
  const { opened, pointAt, startNavigation, releaseNavigation } = await readyEditorWithPendingNavigation();
  let armed = false;
  const image = { pos: 6, rect: { x: 0, y: 0, width: 1, height: 1 } } as unknown as NonNullable<ReturnType<DisplayListQueries['imageAtPoint']>>;
  const spies = [...new Set([opened.harness.renderer.queries!, opened.harness.renderer.inputQueries!])]
    .map((queries) => spyOn(queries, 'imageAtPoint').mockImplementation(() => (armed ? image : null)));
  try {
    await startNavigation();
    const before = opened.session.selection();
    armed = true;
    fireEvent.mouseDown(opened.canvas, pointAt(6));
    armed = false;
    fireEvent.mouseMove(window, pointAt(10));
    fireEvent.mouseUp(window, pointAt(10));
    expect(opened.session.selection()).toEqual(before);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    await act(async () => {
      releaseNavigation();
      await opened.editor.current!.flushPendingInput();
    });
    const anchor = { story: 'body', paraId: opened.session.paragraphs('body')[0].paraId, offset: 5 };
    expect(opened.session.selection()).toEqual({ anchor, head: { ...anchor, offset: 6 } });
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    releaseNavigation();
    for (const spy of spies.reverse()) spy.mockRestore();
    opened.close();
  }
});

test('a composition started during opening commits exactly once after the switch', async () => {
  const opened = await openingEditor();
  const insert = spyOn(opened.preview, 'insertText');
  try {
    const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
    opened.click(6);
    fireEvent.input(textarea, { target: { value: 'A' } });
    fireEvent.compositionStart(textarea);
    fireEvent.compositionUpdate(textarea, { data: '日本' });
    textarea.value = '日本';
    const full = await opened.switchToFull();
    expect(opened.view.getByTestId('yrs-input')).toBe(textarea);
    expect(textarea.value).toBe('日本');
    await opened.presentFull(full);
    fireEvent.compositionEnd(textarea, { data: '日本' });
    fireEvent(textarea, new InputEvent('textInput', { bubbles: true, data: '日本' }));
    fireEvent.input(textarea);
    await act(async () => { await Promise.resolve(); });
    fireEvent.input(textarea);
    const load = spyOn(full, 'loadState');
    try {
      await opened.loadPeer(full);
      expect(full.paragraphs('body')[0].text).toBe('FirstA日本 paragraph');
      expect(full.selection()?.head.offset).toBe(8);
      expect(load).toHaveBeenCalledTimes(1);
      expect(insert).not.toHaveBeenCalled();
      expect(textarea.value).toBe('');
    } finally {
      load.mockRestore();
    }
  } finally {
    insert.mockRestore();
    opened.close();
  }
});

test('input held during opening is discarded when another document replaces it', async () => {
  const opened = await openingEditor(false, false, { delayQueries: true });
  const insert = spyOn(opened.preview, 'insertText');
  try {
    const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
    opened.click(6);
    fireEvent.input(textarea, { target: { value: 'discard' } });
    fireEvent.keyDown(textarea, { key: 'Enter' });
    fireEvent.paste(textarea, { clipboardData: { getData: () => 'old paste' } });
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    const replacement = await opened.replace();
    expect(replacement).not.toBe(opened.preview);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.view.getByTestId('yrs-input')).toBe(textarea);
    fireEvent.input(textarea, { target: { value: 'C' } });
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    opened.releaseInputQueries();
    await opened.loadPeer(full);
    expect(full.paragraphs('body')).toHaveLength(205);
    expect(full.paragraphs('body')[0].text).toBe('CFirst paragraph');
    expect(full.selection()?.head.offset).toBe(1);
    expect(insert).not.toHaveBeenCalled();
    expect(opened.harness.errors).toEqual([]);
    expect(opened.harness.mainOpens).toEqual([false]);
  } finally {
    insert.mockRestore();
    opened.close();
  }
});

test('viewer sessions stay read-only while opening', async () => {
  const opened = await openingViewer();
  const insert = spyOn(opened.session, 'insertText');
  try {
    const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
    const version = opened.session.version();
    expect(textarea.readOnly).toBe(true);
    opened.click();
    fireEvent.input(textarea, { target: { value: 'ignored' } });
    fireEvent.keyDown(textarea, { key: 'Enter' });
    fireEvent.keyDown(textarea, { key: 'Backspace' });
    fireEvent.paste(textarea, { clipboardData: { getData: () => 'ignored paste' } });
    await act(async () => {});
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.session.version()).toBe(version);
    expect(opened.session.storyIds()).toEqual([]);
    const full = await opened.present();
    act(() => opened.harness.pipeline.runLayoutPipeline());
    await waitFor(() => expect(opened.harness.renderer.presentedEngine).toBe(full));
    act(() => opened.harness.presentFrame());
    await opened.frames.settleAndIdle(opened.harness.renderer.settledDisplayList(null, null, 'window'));
    expect((opened.view.getByTestId('yrs-input') as HTMLTextAreaElement).readOnly).toBe(true);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(full.storyIds()).toEqual([]);
    expect(insert).not.toHaveBeenCalled();
    expect(opened.harness.mainOpens).toEqual([]);
  } finally {
    insert.mockRestore();
    opened.close();
  }
});

test('a font announcement while the editor peer is pending replays once after the plugin host opens', async () => {
  const descriptor = Object.getOwnPropertyDescriptor(document, 'fonts');
  const fonts = Object.assign(new EventTarget(), { ready: Promise.resolve() });
  Object.defineProperty(document, 'fonts', { configurable: true, value: fonts });
  const results: unknown[] = [];
  try {
    for (const held of [true, false]) {
      const opened = await pendingFontEditor();
      try {
        if (held) act(() => {
          fonts.dispatchEvent(new Event('loadingdone'));
          fonts.dispatchEvent(new Event('loadingdone'));
        });
        await opened.drain();
        expect(opened.requirements).not.toHaveBeenCalled();
        expect(opened.posted.slice(opened.initializedPosts).filter((request) => request.type === 'fontRequirements'))
          .toHaveLength(0);
        await act(async () => { await opened.tasks.run(); });
        await act(async () => { await opened.tasks.run(); });
        expect(opened.full.paragraphs('body')[0]!.text).toBe('First paragraph');
        expect(replicaHelpers.workerOpenReplicaPending(opened.full)).toBe(true);
        expect(replicaHelpers.workerOpenReplicaHydrating(opened.full)).toBe(false);
        expect(opened.fallback.timers.size).toBe(2);
        if (held) act(() => { fonts.dispatchEvent(new Event('loadingdone')); });
        await opened.drain();
        expect(opened.requirements).not.toHaveBeenCalled();
        expect(opened.posted.slice(opened.initializedPosts).filter((request) => request.type === 'fontRequirements'))
          .toHaveLength(0);
        const beforeLayout = opened.editor()!.getLayout();
        const beforeQueries = opened.readiness.layout().queries;
        const beforePosts = opened.posted.length;
        act(() => opened.readiness.complete(opened.full));
        await act(async () => { await Promise.resolve(); });
        expect(opened.trace).toEqual(['plugin-open']);
        if (!held) act(() => { fonts.dispatchEvent(new Event('loadingdone')); });
        await act(async () => {
          while (opened.tasks.tasks.length) await opened.tasks.run();
        });
        opened.tasks.restore();
        await opened.until(() => opened.editor()!.getLayout() !== beforeLayout &&
          opened.readiness.layout().queries !== beforeQueries && opened.readiness.layout().complete);
        await opened.drain();
        expect(opened.requirements).toHaveBeenCalledTimes(1);
        expect(opened.trace).toEqual(['plugin-open', 'font-refresh']);
        const posts = opened.posted.slice(beforePosts).map((request) => request.type).sort();
        expect(posts.filter((type) => type === 'fontRequirements')).toHaveLength(0);
        expect(posts.filter((type) => type === 'sync')).toHaveLength(1);
        const layout = opened.editor()!.getLayout()!;
        const queries = opened.readiness.layout().queries!;
        expect(queries.isReady()).toBe(true);
        expect(queries.pageCount()).toBe(layout.pages.length);
        expect(sourceVersionOf(queries)).toBe(sourceVersionOf(layout));
        expect(queries.caretRect(1)).not.toBeNull();
        const canvases = [...opened.view.container.querySelectorAll<HTMLElement>('.canvas-page')];
        expect(canvases.length).toBeGreaterThan(0);
        results.push({
          posts,
          pages: layout.pages.map((page) => ({
            size: page.size,
            margins: page.margins,
            fragments: page.fragments.map(({ kind, x, y, width, pmStart, pmEnd }) => ({
              kind, x, y, width, pmStart, pmEnd,
            })),
          })),
          sizes: Array.from({ length: queries.pageCount() }, (_, index) => queries.pageSize(index)),
          caret: queries.caretRect(1),
          canvases: canvases.map((canvas) => ({
            page: canvas.getAttribute('data-page-index'),
            width: canvas.style.width,
            height: canvas.style.height,
          })),
        });
      } finally {
        opened.close();
      }
    }
    expect(results[0]).toEqual(results[1]);
  } finally {
    if (descriptor) Object.defineProperty(document, 'fonts', descriptor);
    else Reflect.deleteProperty(document, 'fonts');
  }
}, 30_000);

test('a held font announcement is dropped when the document is replaced or the editor unmounts before readiness', async () => {
  const descriptor = Object.getOwnPropertyDescriptor(document, 'fonts');
  const fonts = Object.assign(new EventTarget(), { ready: Promise.resolve() });
  Object.defineProperty(document, 'fonts', { configurable: true, value: fonts });
  const results: unknown[] = [];
  try {
    for (const [action, held] of [['replace', true], ['replace', false], ['unmount', true]] as const) {
      const opened = await pendingFontEditor();
      const oldWorker = spyOn(opened.workers[0]!, 'postMessage');
      let restoreReplacementRequirements = () => {};
      let restoreReplacementWorker = () => {};
      try {
        act(() => { fonts.dispatchEvent(new Event('loadingdone')); });
        await act(async () => { await opened.tasks.run(); });
        await act(async () => { await opened.tasks.run(); });
        expect(replicaHelpers.workerOpenReplicaPending(opened.full)).toBe(true);
        expect(opened.requirements).not.toHaveBeenCalled();
        expect(oldWorker.mock.calls.filter(([request]) => request.type === 'fontRequirements')).toHaveLength(0);
        const beforeReplacementPosts = opened.posted.length;
        if (action === 'unmount') opened.view.unmount();
        else {
          let replacement!: Promise<void>;
          act(() => {
            replacement = opened.ref.current!.loadDocumentBuffer(longBytes.slice().buffer as ArrayBuffer);
          });
          await opened.until(() => opened.session() !== null && opened.session() !== opened.full &&
            opened.session()?.isDisplayOnly() === true);
          const fullOpen = () => opened.posted.slice(beforeReplacementPosts).find((request) =>
            request.type === 'open' && request.previewBlocks === undefined);
          await opened.until(() => {
            const request = fullOpen();
            return !!request && opened.replies.has(request.id);
          });
          act(() => opened.reply(fullOpen()!));
          await opened.settle(replacement);
          await opened.until(() => opened.session() !== null && opened.session() !== opened.full &&
            opened.session()?.isDisplayOnly() === false &&
            replicaHelpers.workerOpenReplicaStarted(opened.session()!) &&
            opened.posted.slice(beforeReplacementPosts).some((request) =>
              request.type === 'encodeState' && opened.replies.has(request.id)));
          await opened.drain();
        }
        await act(async () => {
          opened.readiness.complete(opened.full);
          opened.fallback.advance(30_000);
          while (opened.tasks.tasks.length) await opened.tasks.run();
        });
        if (action === 'unmount') opened.tasks.restore();
        await opened.drain();
        expect(opened.requirements).not.toHaveBeenCalled();
        expect(oldWorker.mock.calls.filter(([request]) => request.type === 'fontRequirements')).toHaveLength(0);
        expect(opened.trace).toEqual([]);
        if (action === 'replace') {
          const replacement = opened.session()!;
          const replacementWorker = spyOn(opened.workers.at(-1)!, 'postMessage');
          restoreReplacementWorker = registerRestore(() => replacementWorker.mockRestore());
          const readRequirements = replacement.layoutFontRequirementsJson;
          const requirements = spyOn(replacement, 'layoutFontRequirementsJson').mockImplementation((input) => {
            opened.trace.push('font-refresh:B');
            return readRequirements(input);
          });
          restoreReplacementRequirements = registerRestore(() => requirements.mockRestore());
          const state = opened.posted.slice(beforeReplacementPosts).find((request) => request.type === 'encodeState')!;
          act(() => opened.reply(state));
          await opened.until(() => opened.tasks.tasks.length === 1);
          expect(replacement.isDisplayOnly()).toBe(false);
          expect(replicaHelpers.workerOpenReplicaStarted(replacement)).toBe(true);
          expect(replicaHelpers.workerOpenReplicaPending(replacement)).toBe(true);
          expect(replicaHelpers.workerOpenReplicaHydrating(replacement)).toBe(true);
          expect(requirements).not.toHaveBeenCalled();
          expect(replacementWorker.mock.calls.filter(([request]) => request.type === 'fontRequirements')).toHaveLength(0);
          if (held) act(() => { fonts.dispatchEvent(new Event('loadingdone')); });
          await opened.drain();
          expect(requirements).not.toHaveBeenCalled();
          expect(replacementWorker.mock.calls.filter(([request]) => request.type === 'fontRequirements')).toHaveLength(0);
          await act(async () => { await opened.tasks.run(); });
          await act(async () => { await opened.tasks.run(); });
          expect(replacement.paragraphs('body')[0]!.text).toBe('First paragraph');
          expect(replicaHelpers.workerOpenReplicaPending(replacement)).toBe(true);
          expect(replicaHelpers.workerOpenReplicaHydrating(replacement)).toBe(false);
          expect(opened.fallback.timers.size).toBe(2);
          await opened.drain();
          expect(requirements).not.toHaveBeenCalled();
          expect(replacementWorker.mock.calls.filter(([request]) => request.type === 'fontRequirements')).toHaveLength(0);
          expect(opened.trace).toEqual([]);
          const beforeLayout = opened.editor()!.getLayout();
          const beforeQueries = opened.readiness.layout().queries;
          const beforePosts = opened.posted.length;
          act(() => opened.readiness.complete(replacement));
          await act(async () => { await Promise.resolve(); });
          expect(replicaHelpers.workerOpenReplicaPending(replacement)).toBe(false);
          expect(opened.trace).toEqual(['plugin-open']);
          if (!held) act(() => { fonts.dispatchEvent(new Event('loadingdone')); });
          await act(async () => {
            while (opened.tasks.tasks.length) await opened.tasks.run();
          });
          opened.tasks.restore();
          await opened.until(() => opened.editor()!.getLayout() !== beforeLayout &&
            opened.readiness.layout().queries !== beforeQueries && opened.readiness.layout().complete);
          await opened.drain();
          expect(opened.session()).toBe(replacement);
          expect(opened.requirements).not.toHaveBeenCalled();
          expect(oldWorker.mock.calls.filter(([request]) => request.type === 'fontRequirements')).toHaveLength(0);
          expect(requirements).toHaveBeenCalledTimes(1);
          expect(replacementWorker.mock.calls.filter(([request]) => request.type === 'fontRequirements')).toHaveLength(0);
          expect(opened.trace).toEqual(['plugin-open', 'font-refresh:B']);
          const posts = opened.posted.slice(beforePosts).map((request) => request.type).sort();
          expect(posts.filter((type) => type === 'fontRequirements')).toHaveLength(0);
          expect(posts.filter((type) => type === 'sync')).toHaveLength(1);
          const layout = opened.editor()!.getLayout()!;
          const queries = opened.readiness.layout().queries!;
          expect(queries.isReady()).toBe(true);
          expect(queries.pageCount()).toBe(layout.pages.length);
          expect(sourceVersionOf(queries)).toBe(sourceVersionOf(layout));
          expect(queries.caretRect(1)).not.toBeNull();
          const canvases = [...opened.view.container.querySelectorAll<HTMLElement>('.canvas-page')];
          expect(canvases.length).toBeGreaterThan(0);
          results.push({
            posts,
            pages: layout.pages.map((page) => ({
              size: page.size,
              margins: page.margins,
              fragments: page.fragments.map(({ kind, x, y, width, pmStart, pmEnd }) => ({
                kind, x, y, width, pmStart, pmEnd,
              })),
            })),
            sizes: Array.from({ length: queries.pageCount() }, (_, index) => queries.pageSize(index)),
            caret: queries.caretRect(1),
            canvases: canvases.map((canvas) => ({
              page: canvas.getAttribute('data-page-index'),
              width: canvas.style.width,
              height: canvas.style.height,
            })),
          });
        }
      } finally {
        restoreReplacementRequirements();
        restoreReplacementWorker();
        oldWorker.mockRestore();
        opened.close();
      }
    }
    expect(results).toHaveLength(2);
    expect(results[0]).toEqual(results[1]);
  } finally {
    if (descriptor) Object.defineProperty(document, 'fonts', descriptor);
    else Reflect.deleteProperty(document, 'fonts');
  }
}, 30_000);

test('font announcements in viewer sessions are not held', async () => {
  const descriptor = Object.getOwnPropertyDescriptor(document, 'fonts');
  const fonts = Object.assign(new EventTarget(), { ready: Promise.resolve() });
  Object.defineProperty(document, 'fonts', { configurable: true, value: fonts });
  let opened: Awaited<ReturnType<typeof openingViewer>> | undefined;
  let restoreReplicaState = () => {};
  let restoreRequirements = () => {};
  try {
    opened = await openingViewer();
    const full = await opened.present();
    await opened.frames.settleAndIdle(opened.harness.renderer.settledDisplayList(null, null, 'window'));
    await act(async () => {});
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    const started = replicaHelpers.workerOpenReplicaStarted;
    const pending = replicaHelpers.workerOpenReplicaPending;
    const forcedStarted = spyOn(replicaHelpers, 'workerOpenReplicaStarted')
      .mockImplementation((session) => session === full || started(session));
    const forcedPending = spyOn(replicaHelpers, 'workerOpenReplicaPending')
      .mockImplementation((session) => session === full || pending(session));
    restoreReplicaState = registerRestore(() => {
      forcedPending.mockRestore();
      forcedStarted.mockRestore();
    });
    const beforePosts = opened.posted.length;
    const beforeLayout = opened.editor.current!.getLayout();
    const requirements = spyOn(full, 'layoutFontRequirementsJson');
    restoreRequirements = registerRestore(() => requirements.mockRestore());
    expect(workerOpenDocumentHeld(full)).toBe(true);
    expect(opened.harness.core.experimentalWorkerOpen).toBe(true);
    expect(opened.harness.core.previewing).toBe(false);
    expect(opened.editor.current!.isWorkerViewer()).toBe(true);
    expect(full.isDisplayOnly()).toBe(false);
    expect(replicaHelpers.workerOpenReplicaStarted(full)).toBe(true);
    expect(replicaHelpers.workerOpenReplicaPending(full)).toBe(true);
    expect(registeredWorkerProposalAuthority(full)?.holdsWorkerState() ?? false).toBe(false);
    expect(opened.harness.core.replicaReady).toBe(false);
    act(() => { fonts.dispatchEvent(new Event('loadingdone')); });
    await waitFor(() => expect(opened!.posted.slice(beforePosts)
      .filter((request) => request.type === 'fontRequirements')).toHaveLength(1));
    await waitFor(() => expect(opened!.editor.current!.getLayout()).not.toBe(beforeLayout));
    expect(opened.posted.slice(beforePosts).filter((request) => request.type === 'sync')).toHaveLength(1);
    expect(opened.posted.slice(beforePosts).filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(requirements).not.toHaveBeenCalled();
    expect(replicaHelpers.workerOpenReplicaPending(full)).toBe(true);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.harness.mainOpens).toEqual([]);
    expect(full.storyIds()).toEqual([]);
  } finally {
    try {
      opened?.close();
    } finally {
      restoreRequirements();
      restoreReplicaState();
      if (descriptor) Object.defineProperty(document, 'fonts', descriptor);
      else Reflect.deleteProperty(document, 'fonts');
    }
  }
});

test('held input and loading reads wait for the current worker layout after hydration', async () => {
  const opened = await openingEditor(true, false, {
    holdFullLayout: true,
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    opened.click(6);
    opened.type('A');
    fireEvent.keyDown(opened.view.getByTestId('yrs-input'), { key: 'ArrowLeft' });
    opened.type('B');
    const pending = requestWorkerOpenReplica(full)!;
    let ready = false;
    void pending.then(() => { ready = true; });
    await finishHeldHydration(opened, tasks);
    const completion = await opened.received('completeLayout');
    expect(opened.replies.has(completion.id)).toBe(true);
    expect(full.paragraphs('body')[0].text).toBe('First paragraph');
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.harness.core.replicaReadyRef?.current).toBe(false);
    expect(ready).toBe(false);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(opened.harness.ref.current!.getDocument()).toBeNull();
    expect(opened.harness.ref.current!.findInDocument('First')).toEqual([]);
    expect(() => opened.harness.ref.current!.insertBreak({ paraId: '00000001', type: 'page' })).toThrow(DocxReplicaNotReadyError);
    let saved = false;
    const save = opened.harness.ref.current!.save().then(() => { saved = true; });
    await act(async () => {});
    expect(saved).toBe(false);
    await act(async () => {
      opened.reply(completion);
      await pending;
    });
    expect(opened.harness.core.replicaReady).toBe(true);
    await act(async () => opened.editor.current!.flushPendingInput());
    await save;
    expect(full.paragraphs('body')[0].text).toBe('FirstBA paragraph');
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.ref.current!.getDocument()).not.toBeNull();
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('a completed worker layout commits readiness as soon as hydration finishes', async () => {
  const opened = await openingEditor();
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    await finishHeldHydration(opened, tasks);
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(opened.harness.core.replicaReadyRef?.current).toBe(true);
    expect(fallback.timers.size).toBe(0);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('a first worker pass that completes layout commits readiness and replays held input without a completion request', async () => {
  const opened = await openingEditor();
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full);
    opened.click(6);
    opened.type('A');
    fireEvent.keyDown(opened.view.getByTestId('yrs-input'), { key: 'ArrowLeft' });
    opened.type('B');
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(opened.harness.core.replicaReady).toBe(false);
    await finishHeldHydration(opened, tasks);
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(opened.harness.core.replicaReadyRef?.current).toBe(true);
    await waitFor(() => {
      expect(full.paragraphs('body')[0].text).toBe('FirstBA paragraph');
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
    });
    expect(opened.posted.some((request) => request.type === 'completeLayout')).toBe(false);
    expect(fallback.timers.size).toBe(0);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('a new provisional layout keeps held input pending after an earlier pass completes', async () => {
  const opened = await openingEditor(true, false, {
    holdReply: (request) => request.type === 'completeLayout',
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    opened.click(6);
    opened.type('Q');
    const pending = requestWorkerOpenReplica(full)!;
    let ready = false;
    void pending.then(() => { ready = true; });
    act(() => {
      replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'complete');
      replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'provisional');
    });
    await finishHeldHydration(opened, tasks);
    expect(ready).toBe(false);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.harness.core.replicaReadyRef?.current).toBe(false);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(full.paragraphs('body')[0].text).toBe('First paragraph');
    expect(fallback.timers.size).toBe(2);
    await act(async () => {
      replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'complete');
      await pending;
    });
    expect(ready).toBe(true);
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(opened.harness.core.replicaReadyRef?.current).toBe(true);
    await waitFor(() => {
      expect(full.paragraphs('body')[0].text).toBe('FirstQ paragraph');
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
    });
    expect(fallback.timers.size).toBe(0);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('page progress after layout completion commits readiness as soon as hydration finishes', async () => {
  const opened = await openingEditor(true, false, {
    holdReply: (request) => request.type === 'completeLayout',
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    const pending = requestWorkerOpenReplica(full)!;
    let ready = false;
    void pending.then(() => { ready = true; });
    act(() => {
      replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'complete');
      replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'page');
    });
    expect(ready).toBe(false);
    await finishHeldHydration(opened, tasks);
    expect(ready).toBe(true);
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(opened.harness.core.replicaReadyRef?.current).toBe(true);
    expect(fallback.timers.size).toBe(0);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('flushing held input commits a hydrated peer before worker layout completes', async () => {
  const opened = await openingEditor(true, false, {
    holdFullLayout: true,
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    opened.click(6);
    opened.type('Q');
    requestWorkerOpenReplica(full);
    await finishHeldHydration(opened, tasks);
    const completion = await opened.received('completeLayout');
    expect(opened.harness.core.replicaReady).toBe(false);
    await act(async () => opened.editor.current!.flushPendingInput());
    expect(opened.replies.has(completion.id)).toBe(true);
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(full.paragraphs('body')[0].text).toBe('FirstQ paragraph');
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('an editor readiness request commits a hydrated peer before worker layout completes', async () => {
  const opened = await openingEditor(true, false, {
    holdFullLayout: true,
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    requestWorkerOpenReplica(full);
    await finishHeldHydration(opened, tasks);
    const completion = await opened.received('completeLayout');
    expect(opened.harness.core.replicaReady).toBe(false);
    act(() => replicaHelpers.requestWorkerOpenReplicaReadiness(full));
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(opened.replies.has(completion.id)).toBe(true);
    expect(fallback.timers.size).toBe(0);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('layout progress keeps readiness pending past ten seconds until it stalls for three seconds', async () => {
  const opened = await openingEditor(true, false, {
    holdFullLayout: true,
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    requestWorkerOpenReplica(full);
    await finishHeldHydration(opened, tasks);
    for (let step = 0; step < 6; step += 1) {
      act(() => {
        fallback.advance(2000);
        replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'page');
      });
      expect(opened.harness.core.replicaReady).toBe(false);
    }
    act(() => fallback.advance(2999));
    expect(opened.harness.core.replicaReady).toBe(false);
    act(() => fallback.advance(1));
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(fallback.timers.size).toBe(0);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('readiness commits thirty seconds after hydration even with ongoing layout progress', async () => {
  const opened = await openingEditor(true, false, {
    holdFullLayout: true,
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    requestWorkerOpenReplica(full);
    await finishHeldHydration(opened, tasks);
    for (let step = 0; step < 14; step += 1) {
      act(() => {
        fallback.advance(2000);
        replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'provisional');
      });
      expect(opened.harness.core.replicaReady).toBe(false);
    }
    act(() => fallback.advance(1999));
    expect(opened.harness.core.replicaReady).toBe(false);
    act(() => fallback.advance(1));
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(fallback.timers.size).toBe(0);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('replacing a document discards its hydrated peer waiting for layout completion', async () => {
  const opened = await openingEditor(true, false, {
    holdFullLayout: true,
  });
  const tasks = holdHydrationTasks();
  const fallback = holdLayoutFallback();
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    opened.click(6);
    opened.type('Q');
    const pending = requestWorkerOpenReplica(full)!;
    await finishHeldHydration(opened, tasks);
    const completion = await opened.received('completeLayout');
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(fallback.timers.size).toBe(2);
    await opened.replace();
    await expect(pending).rejects.toThrow('The document changed');
    expect(fallback.timers.size).toBe(0);
    await act(async () => {
      opened.reply(completion);
      replicaHelpers.notifyWorkerOpenLayoutProgress(full, 'complete');
      fallback.advance(30_000);
    });
    const replacement = await opened.switchToFull();
    expect(replacement).not.toBe(full);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.harness.core.replicaReadyRef?.current).toBe(false);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    fallback.restore();
    tasks.restore();
  }
});

test('worker hydration opens and loads in separate tasks before publishing readiness', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  try {
    const { result } = renderHook(useHarness, { initialProps: {
      ...initialProps,
      collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    } });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const load = spyOn(session, 'loadState');
    restoreLoad = registerRestore(() => load.mockRestore());
    let ready = false;
    const pending = requestWorkerOpenReplica(session)!;
    void pending.then(() => { ready = true; });
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    expect(result.current.mainOpens).toEqual([false]);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.core.replicaReadyRef?.current).toBe(false);
    expect(ready).toBe(false);
    expect(replicas).toEqual([]);
    await act(async () => tasks.run());
    expect(load).toHaveBeenCalledTimes(1);
    expect(result.current.core.replicaReady).toBe(false);
    expect(result.current.core.replicaReadyRef?.current).toBe(false);
    expect(ready).toBe(false);
    expect(replicas).toEqual([]);
    await act(async () => {
      replicaHelpers.notifyWorkerOpenLayoutProgress(session, 'complete');
      await tasks.run();
      await pending;
    });
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.core.replicaReadyRef?.current).toBe(true);
    expect(ready).toBe(true);
    expect(replicas).toEqual([session]);
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test.each(['replace', 'unmount'] as const)('a document %s between hydration tasks stops the stale peer', async (action) => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  try {
    const props = { ...initialProps,
      collaboration: { onReplica: (session: unknown) => replicas.push(session as YrsSession | null) },
    };
    const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const errors = result.current.errors;
    const load = spyOn(session, 'loadState');
    restoreLoad = registerRestore(() => load.mockRestore());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    expect(result.current.mainOpens).toEqual([false]);
    if (action === 'replace') {
      rerender({ ...props, source: bytes.slice(), generation: 2 });
      await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
      expect(result.current.core.session).not.toBe(session);
    } else {
      unmount();
    }
    await act(async () => tasks.run());
    await expect(pending).rejects.toThrow('The document changed');
    expect(load).not.toHaveBeenCalled();
    expect(replicas).toEqual([]);
    expect(errors).toEqual([]);
    if (action === 'replace') {
      expect(result.current.mainOpens).toEqual([false]);
      expect(result.current.core.replicaReady).toBe(false);
      expect(result.current.core.replicaReadyRef?.current).toBe(false);
    }
  } finally {
    try {
      cleanup();
    } finally {
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test('overlay projection builds in a task after replica readiness', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const fonts = Object.getOwnPropertyDescriptor(document, 'fonts');
  const restoreFonts = registerRestore(() => {
    if (fonts) Object.defineProperty(document, 'fonts', fonts);
    else Reflect.deleteProperty(document, 'fonts');
  });
  let restoreSegments = () => {};
  try {
    if (!document.fonts) Object.defineProperty(document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
    });
    const { result } = renderHook(useHarness, { initialProps });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.queries?.isReady()).toBe(true));
    const segments = spyOn(session, 'storySegments');
    restoreSegments = registerRestore(() => segments.mockRestore());
    const target = document.createElement('div');
    const host = createRef<HTMLDivElement>();
    const editor = () => <PagedEditor document={result.current.host!.document}
      yrsCore={result.current.core} readOnly holdInput
      canvasOverlayTarget={target} canvasHostRef={host}
      displayListQueries={result.current.renderer.queries}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
      fontRequirementsInWorker={result.current.renderer.fontRequirementsInWorker}
      layoutInWorker={result.current.renderer.layoutInWorker} />;
    const view = render(editor());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    await act(async () => tasks.run());
    expect(segments).not.toHaveBeenCalled();
    expect(result.current.core.replicaReady).toBe(false);
    await act(async () => {
      await tasks.run();
      await pending;
    });
    view.rerender(editor());
    expect(result.current.core.replicaReady).toBe(true);
    expect(segments).not.toHaveBeenCalled();
    expect(tasks.tasks).toHaveLength(1);
    await act(async () => tasks.run());
    expect(segments).toHaveBeenCalled();
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      restoreSegments();
      restoreFonts();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test.each([1, 2])('synchronous ensure finishes the worker peer at hydration yield %s once', async (boundary) => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  try {
    const { result } = renderHook(useHarness, { initialProps: {
      ...initialProps,
      collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    } });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const load = spyOn(session, 'loadState');
    restoreLoad = registerRestore(() => load.mockRestore());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    if (boundary === 2) await act(async () => tasks.run());
    act(() => {
      ensureWorkerOpenReplica(session);
      expect(session.hasStory('body')).toBe(true);
      expect(result.current.core.replicaReadyRef?.current).toBe(true);
      expect(load).toHaveBeenCalledTimes(1);
    });
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(replicas).toEqual([session]);
    await act(async () => {
      await tasks.run();
      await pending;
    });
    expect(load).toHaveBeenCalledTimes(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(replicas).toEqual([session]);
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test.each([1, 2])('inactive editors keep base synchronous hydration and no proposal authority at yield %s through save and worker loss', async (boundary) => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  try {
    const { result } = renderHook(useHarness, { initialProps: {
      ...initialProps,
      collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    } });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    expect(registeredWorkerProposalAuthority(session)).toBeNull();
    const load = spyOn(session, 'loadState');
    restoreLoad = registerRestore(() => load.mockRestore());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    if (boundary === 2) await act(async () => tasks.run());
    act(() => {
      expect(registeredWorkerProposalAuthority(session)).toBeNull();
      ensureWorkerOpenReplica(session);
      expect(registeredWorkerProposalAuthority(session)).toBeNull();
      expect(session.hasStory('body')).toBe(true);
      expect(result.current.core.replicaReadyRef?.current).toBe(true);
      expect(load).toHaveBeenCalledTimes(1);
    });
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(replicas).toEqual([session]);
    await act(async () => {
      await tasks.run();
      await pending;
    });
    expect(load).toHaveBeenCalledTimes(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(replicas).toEqual([session]);
    expect(registeredWorkerProposalAuthority(session)).toBeNull();
    await act(async () => {
      expect(await workerOpenSave(session)!.save([], session)).toBeInstanceOf(ArrayBuffer);
    });
    expect(registeredWorkerProposalAuthority(session)).toBeNull();
    spyOn(console, 'error').mockImplementation(() => {});
    await act(async () => {
      workers[0]!.onerror?.({ message: 'inactive editor worker crashed' } as ErrorEvent);
      expect(await result.current.renderer.layoutInWorker(session, JSON.stringify({
        bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
        measurement: { defaults: { fontFamily: 'Calibri', fontSize: 11 } }, renderEnv: {},
      }))).toBeNull();
    });
    expect(registeredWorkerProposalAuthority(session)).toBeNull();
    expect(result.current.core.replicaReadyRef?.current).toBe(true);
    expect(result.current.core.replicaReady).toBe(true);
    expect(replicas).toEqual([session]);
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test.each([false, true])('a loadState error after yielding preserves replica fallback with failure=%s', async (fails) => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const replicas: Array<YrsSession | null> = [];
  let restoreLoad = () => {};
  let restoreOpen = () => {};
  try {
    const { result } = renderHook(useHarness, { initialProps: {
      ...initialProps,
      collaboration: { onReplica: (session) => replicas.push(session as YrsSession | null) },
    } });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const loadError = new Error('State load failed');
    const fallbackError = new Error('Replica fallback failed');
    const load = spyOn(session, 'loadState').mockImplementation(() => { throw loadError; });
    restoreLoad = registerRestore(() => load.mockRestore());
    const originalOpen = session.openDocx.bind(session);
    const fallbackSeeds: boolean[] = [];
    const open = spyOn(session, 'openDocx').mockImplementation((source, seed, options) => {
      fallbackSeeds.push(seed);
      if (fails && seed) throw fallbackError;
      return originalOpen(source, seed, options);
    });
    restoreOpen = registerRestore(() => open.mockRestore());
    const pending = requestWorkerOpenReplica(session)!;
    await waitFor(() => expect(posted.some((request) => request.type === 'encodeState')).toBe(true));
    await act(async () => workers[0].release());
    await waitFor(() => expect(tasks.tasks).toHaveLength(1));
    expect(load).not.toHaveBeenCalled();
    expect(result.current.errors).toEqual([]);
    await act(async () => {
      await tasks.run();
      if (fails) await expect(pending).rejects.toBe(fallbackError);
      else await pending;
    });
    expect(load).toHaveBeenCalledTimes(1);
    expect(fallbackSeeds).toEqual([false, true]);
    expect(result.current.core.replicaReady).toBe(!fails);
    expect(result.current.core.replicaReadyRef?.current).toBe(!fails);
    expect(replicas).toEqual(fails ? [] : [session]);
    expect(result.current.errors).toEqual(fails ? [fallbackError] : []);
  } finally {
    try {
      cleanup();
    } finally {
      restoreOpen();
      restoreLoad();
      visibility.restore();
      tasks.restore();
      frames.restore();
      globalThis.Worker = originalWorker;
    }
  }
});

test('a hidden document starts the editor peer immediately without layout or idle', async () => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const fallback = holdLayoutFallback();
  const visibility = stubDocumentVisibility('hidden');
  const { result, unmount } = renderHook(useHarness, { initialProps });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    expect(result.current.renderer.frame).toBeNull();
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(true);
    expect(result.current.mainOpens).toEqual([]);
    act(() => frames.run());
    act(() => frames.run());
    act(() => frames.runIdle());
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    await act(async () => workers[0].release());
    await waitFor(() => expect(fallback.timers.size).toBe(2));
    expect(result.current.core.replicaReady).toBe(false);
    await act(async () => {
      fallback.advance(3000);
      await awaitWorkerOpenReplica(session);
    });
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    visibility.restore();
    fallback.restore();
    frames.restore();
  }
});

test.each([false, true])('hiding a tab starts the editor peer immediately with settledLayout=%s', async (settledLayout) => {
  const { workers, posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const layoutReady = settledLayout ? undefined : new Promise<void>(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, layoutReady },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.status).toBe('ready'));
    act(() => result.current.presentFrame());
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    act(() => visibility.hide());
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(true);
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    act(() => frames.run());
    act(() => frames.run());
    act(() => frames.runIdle());
    act(() => visibility.hide());
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    visibility.restore();
    frames.restore();
  }
});

test('a viewer skips peer scheduling and defers background page builds until idle', async () => {
  const source = await longFixture(1200);
  const options = { holdCompletion: true };
  const { posted, workers } = installWorker(options);
  const frames = holdFrames(true);
  const fallback = holdPeerFallback();
  const visibility = stubDocumentVisibility('visible');
  const onLayoutWait = mock(() => {});
  const waitForUpdate = async (assertion: () => void) => {
    const deadline = performance.now() + 5000;
    for (;;) {
      await act(async () => {
        await new Promise<void>((resolve) => setImmediate(resolve));
      });
      try {
        assertion();
        return;
      } catch (error) {
        if (performance.now() >= deadline) throw error;
      }
    }
  };
  try {
    const { result } = renderHook(useHarness, {
      initialProps: { ...initialProps, source, readOnly: true, viewer: true, onLayoutWait },
    });
    await waitForUpdate(() => expect(result.current.host).not.toBeNull());
    await waitForUpdate(() => expect(result.current.renderer.frame).not.toBeNull());
    expect(result.current.renderer.presentedEngine).toBe(result.current.core.session);
    act(() => result.current.presentFrame());
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(fallback.timers.size).toBe(0);
    await waitForUpdate(() => expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1));
    await waitForUpdate(() => expect(posted.some((request) => request.type === 'completeLayout')).toBe(true));
    options.holdCompletion = false;
    await act(async () => workers[0].release());
    await waitForUpdate(() => expect(result.current.renderer.displayList!.pages.length).toBeGreaterThan(7));
    await waitForUpdate(() => expect(frames.idleCallbacks.size).toBeGreaterThan(1));
    expect(result.current.renderer.displayList!.pages.slice(5, 7).every((page) => page.unbuilt)).toBe(true);
    expect(posted.filter((request) =>
      request.type === 'buildPages' && request.background === true
    )).toEqual([]);
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(fallback.timers.size).toBe(0);
    expect([...frames.idleCallbacks.values()].some(({ options }) => options?.timeout === 2000)).toBe(false);
    await act(async () => frames.runIdle());
    await waitForUpdate(() => expect(posted.some((request) =>
      request.type === 'buildPages' && request.background === true
    )).toBe(true));
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(posted.filter((request) => request.type === 'revisionCount')).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      visibility.restore();
      fallback.restore();
      frames.restore();
    }
  }
}, 15_000);

test('the editor state prefetch waits for a sliced background page-build reply', async () => {
  const options = {
    holdCompletion: true,
    holdReply: (request: ResidentEngineWorkerRequest) =>
      request.type === 'buildPages' || request.type === 'encodeState',
  };
  const { posted, workers, received, reply, replies } = installWorker(options);
  const frames = holdFrames(true);
  const fallback = holdPeerFallback();
  const visibility = stubDocumentVisibility('visible');
  try {
    const props = { ...initialProps, source: await longFixture(1200) };
    const { result, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    await waitFor(() => expect(posted.some((request) => request.type === 'completeLayout')).toBe(true));
    act(() => result.current.presentFrame());
    const layout = result.current.renderer.settledDisplayList(null, null, 'document');
    void layout.catch(() => {});
    options.holdCompletion = false;
    await act(async () => workers[0].release());
    let batch = await received('buildPages');
    while (batch.type === 'buildPages' && !batch.background) {
      await act(async () => reply(batch));
      batch = await received('buildPages', batch.id);
    }
    if (batch.type !== 'buildPages') throw new Error('expected a page build');
    expect(batch.background).toBe(true);
    expect(batch.pages.length).toBeGreaterThan(4);
    expect(batch.pages).toEqual(expect.arrayContaining([5, 6]));
    expect(replies.has(batch.id)).toBe(true);
    const window = result.current.renderer.settledDisplayList(null, null, 'window');
    void window.catch(() => {});
    await act(async () => {
      frames.run();
      frames.runIdle();
    });
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    const postedBeforeReply = posted.length;
    await act(async () => reply(batch));
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    const prefetched = await received('encodeState');
    expect(posted.indexOf(prefetched)).toBeGreaterThanOrEqual(postedBeforeReply);
    expectStatePrefetchAfterPageBuilds(posted);
    await frames.untilCommitted(layout);
    await window;
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    await act(async () => {
      reply(prefetched);
      frames.runIdle();
    });
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expectStatePrefetchAfterPageBuilds(posted);
    expect(load).toHaveBeenCalledTimes(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    visibility.restore();
    fallback.restore();
    frames.restore();
  }
}, 15_000);

test('the editor peer waits for window layout to settle and then starts on idle', async () => {
  let holdMargin = true;
  let marginId: number | null = null;
  const options = {
    holdCompletion: true,
    holdReply: (request: ResidentEngineWorkerRequest) => {
      if (!holdMargin || request.type !== 'buildPages' || !request.background) return false;
      marginId ??= request.id;
      return request.id === marginId;
    },
  };
  const { posted, workers, received, reply } = installWorker(options);
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  try {
    const props = { ...initialProps, source: await longFixture(1200) };
    const { result, unmount } = renderHook(useHarness, { initialProps: props });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const full = result.current.core.session!;
    const load = spyOn(full, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(full));
    await waitFor(() => expect(posted.some((request) => request.type === 'completeLayout')).toBe(true));
    act(() => result.current.presentFrame());
    let settled = false;
    const layout = result.current.renderer.settledDisplayList(null, null, 'window');
    void layout.then(() => { settled = true; }, () => {});
    await act(async () => {
      frames.run();
      frames.runIdle();
    });
    expect(settled).toBe(false);
    expect([...frames.idleCallbacks.values()].some(({ options }) => options?.timeout === 2000)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.core.replicaReady).toBe(false);
    options.holdCompletion = false;
    await act(async () => workers[0].release());
    await waitFor(() => expect(posted.some((request) => request.type === 'buildPages')).toBe(true));
    await waitFor(() => expect(marginId).not.toBeNull());
    const margin = await received('buildPages', marginId! - 1);
    expect(margin).toMatchObject({ type: 'buildPages', pages: [5, 6], background: true });
    expect(settled).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(0);
    holdMargin = false;
    await act(async () => reply(margin));
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(posted.findIndex((request) => request.type === 'encodeState')).toBeGreaterThan(posted.indexOf(margin));
    await frames.untilCommitted(layout);
    expect(settled).toBe(true);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect([...frames.idleCallbacks.values()].some(({ options }) => options?.timeout === 2000)).toBe(true);
    await act(async () => frames.runIdle());
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(posted.some((request) => request.type === 'encodeState')).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
    unmount();
  } finally {
    cleanup();
    visibility.restore();
    frames.restore();
  }
}, 15_000);

test.each(['resolved', 'in flight'] as const)(
  'a host change after state prefetch is %s hydrates the current worker state',
  async (stage) => {
    const worker = installWorker({ holdReply: (request) => request.type === 'encodeState' });
    const frames = holdFrames(true);
    const visibility = stubDocumentVisibility('visible');
    let resident!: NonNullable<Awaited<ReturnType<OpenInWorker>>>;
    const openInWorker: OpenInWorker = async (...args) => {
      const opened = await result.current.renderer.openInWorker(...args);
      if (opened) resident = opened;
      return opened;
    };
    const { result, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, source: await longFixture(1), openInWorker },
    });
    try {
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const session = result.current.core.session!;
      const load = spyOn(session, 'loadState');
      registerRestore(() => load.mockRestore());
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
      act(() => result.current.presentFrame());
      await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
      const prefetched = await worker.received('encodeState');
      const before = worker.responses.get(prefetched)!;
      if (!before.ok || !before.state) throw new Error('expected prefetched state');
      expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
      expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
      expect(result.current.mainOpens).toEqual([]);
      expect(load).not.toHaveBeenCalled();
      if (stage === 'resolved') await act(async () => worker.reply(prefetched));
      const identities = await resident.documentRead({ kind: 'paragraphIdentities' });
      const paragraph = identities.value.paragraphs.find((entry) => entry.session?.story === 'body')!.session!;
      await act(async () => {
        const changed = await resident.proposal({
          kind: 'propose',
          request: {
            expectVersion: identities.version,
            proposals: [{
              id: 'after-prefetch', paragraph,
              suggest: { author: 'Host', date: '2026-10-04T00:00:00Z' },
              op: 'insertText', at: 'start', text: 'Changed ',
            }],
          },
        });
        expect(changed.result).toMatchObject({ ok: true });
      });
      expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
      await act(async () => frames.runIdle());
      await waitFor(() => expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(2));
      const fresh = await worker.received('encodeState', prefetched.id);
      const after = worker.responses.get(fresh)!;
      if (!after.ok || !after.state) throw new Error('expected current worker state');
      expect(after.version).not.toBe(before.version);
      expect(load).not.toHaveBeenCalled();
      await act(async () => {
        if (stage === 'in flight') worker.reply(prefetched);
        worker.reply(fresh);
        await awaitWorkerOpenReplica(session);
      });
      expect(load).toHaveBeenCalledTimes(1);
      expect(load.mock.calls[0]![0]).toEqual(new Uint8Array(after.state));
      expect(load.mock.calls[0]![0]).not.toEqual(new Uint8Array(before.state));
      expect(session.paragraphs('body')[0].text).toBe('Changed First paragraph');
      expect(result.current.mainOpens).toEqual([false]);
      expect(result.current.core.replicaReady).toBe(true);
      expect(result.current.errors).toEqual([]);
    } finally {
      unmount();
      visibility.restore();
      frames.restore();
    }
  },
  15_000
);

test('an already settled editor window prefetches state when its peer begins waiting', async () => {
  const { posted, received, reply } = installWorker({ holdReply: (request) => request.type === 'encodeState' });
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const onLayoutWait = mock(() => {});
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, source: await longFixture(1), onLayoutWait },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    act(() => result.current.presentFrame());
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
    const prefetched = await received('encodeState');
    expect(onLayoutWait).toHaveBeenCalledTimes(1);
    expect(posted.some((request) => request.type === 'buildPages')).toBe(false);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    await act(async () => reply(prefetched));
    expect(load).not.toHaveBeenCalled();
    await act(async () => frames.runIdle());
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(load).toHaveBeenCalledTimes(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    visibility.restore();
    frames.restore();
  }
});

test('a failed state prefetch takes the encode-failure fallback without encoding again', async () => {
  const { workers, posted, received, reply, replies } = installWorker({
    holdReply: (request) => request.type === 'encodeState',
  });
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const { result, unmount } = renderHook(useHarness, {
    initialProps: { ...initialProps, source: await longFixture(1) },
  });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    act(() => result.current.presentFrame());
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
    const prefetched = await received('encodeState');
    const lateState = replies.get(prefetched.id)!;
    await act(async () => {
      workers[0].onmessage?.({
        data: { id: prefetched.id, ok: false, error: 'prefetch failed' },
      } as MessageEvent);
      reply(prefetched);
    });
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.errors).toEqual([]);
    await act(async () => frames.runIdle());
    await waitFor(() => expect(result.current.mainOpens).toEqual([true]));
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(result.current.errors).toEqual([]);
    await act(async () => lateState());
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(load).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([true]);
  } finally {
    unmount();
    visibility.restore();
    frames.restore();
  }
});

test.each([
  ['replace', 'resolved'], ['replace', 'in flight'],
  ['unmount', 'resolved'], ['unmount', 'in flight'],
] as const)(
  'a document %s discards a %s state prefetch before peer start',
  async (action, stage) => {
    let holdStateReply = true;
    const { posted, received, reply, responses } = installWorker({
      holdReply: (request) => holdStateReply && request.type === 'encodeState',
    });
    const frames = holdFrames(true);
    const visibility = stubDocumentVisibility('visible');
    const { result, rerender, unmount } = renderHook(useHarness, {
      initialProps: { ...initialProps, source: await longFixture(1) },
    });
    try {
      await waitFor(() => expect(result.current.host).not.toBeNull());
      const previous = result.current.core.session!;
      const loadPrevious = spyOn(previous, 'loadState');
      registerRestore(() => loadPrevious.mockRestore());
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(previous));
      act(() => result.current.presentFrame());
      await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
      const prefetched = await received('encodeState');
      const oldState = responses.get(prefetched)!;
      if (!oldState.ok || !oldState.state) throw new Error('expected previous document state');
      if (stage === 'resolved') await act(async () => reply(prefetched));
      expect(replicaHelpers.workerOpenReplicaStarted(previous)).toBe(false);
      const staleIdle = [...frames.idleCallbacks.values()].filter(({ options }) => options?.timeout === 2000);
      expect(staleIdle).toHaveLength(1);
      holdStateReply = false;
      if (action === 'replace') {
        const source = await longFixture(2);
        act(() => rerender({ ...initialProps, source, generation: 2 }));
        await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
      } else unmount();
      await act(async () => {
        if (stage === 'in flight') reply(prefetched);
        for (const { callback } of staleIdle) callback({ didTimeout: false, timeRemaining: () => 50 });
        frames.runIdle();
      });
      expect(loadPrevious).not.toHaveBeenCalled();
      expect(replicaHelpers.workerOpenReplicaPending(previous)).toBe(false);
      expect(result.current.mainOpens).toEqual([]);
      if (action === 'replace') {
        const replacement = result.current.core.session!;
        const load = spyOn(replacement, 'loadState');
        registerRestore(() => load.mockRestore());
        act(() => result.current.pipeline.runLayoutPipeline());
        await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(replacement));
        act(() => result.current.presentFrame());
        await frames.settleAndIdle(result.current.renderer.settledDisplayList(null, null, 'window'));
        await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
        expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(2);
        expect(load).toHaveBeenCalledTimes(1);
        expect(load.mock.calls[0]![0]).not.toEqual(new Uint8Array(oldState.state));
        expect(texts(replacement).body).toEqual(['First paragraph', 'Tail paragraph']);
        expect(loadPrevious).not.toHaveBeenCalled();
        expect(result.current.mainOpens).toEqual([false]);
      } else expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
      expect(result.current.errors).toEqual([]);
    } finally {
      if (action === 'replace') unmount();
      visibility.restore();
      frames.restore();
    }
  },
  15_000
);

function observePluginFallback<T>(promise: Promise<T>, clock: ReturnType<typeof holdPeerFallback>) {
  const outcome = {
    settled: false,
    value: undefined as T | undefined,
    error: undefined as unknown,
    at: undefined as number | undefined,
  };
  void promise.then(
    (value) => { outcome.settled = true; outcome.value = value; outcome.at = clock.now; },
    (error: unknown) => { outcome.settled = true; outcome.error = error; outcome.at = clock.now; }
  );
  return outcome;
}

async function flushPluginFallback(clock: ReturnType<typeof holdPeerFallback>) {
  for (let turn = 0; turn < 12; turn += 1) {
    await act(async () => {
      clock.advance(0);
      for (let microtask = 0; microtask < 12; microtask += 1) await Promise.resolve();
    });
  }
}

async function openingPluginFallback() {
  const source = await longFixture(2);
  const worker = installWorker({ holdState: true, holdCompletion: true });
  const frames = holdFrames(true);
  const clock = holdPeerFallback({ allTimers: true });
  const tasks = holdHydrationTasks();
  const visibility = stubDocumentVisibility('visible');
  const deferred = spyOn(replicaHelpers, 'deferWorkerOpenReplica');
  let opened!: NonNullable<Awaited<ReturnType<OpenInWorker>>>;
  const { result, unmount } = renderHook(useHarness, {
    initialProps: {
      ...initialProps, source, layoutCompleteSession: null,
      onWorkerOpen: (worker) => { if (worker) opened = worker; },
    },
  });
  for (let turn = 0; turn < 1_000 && (!opened || result.current.host === null); turn += 1) {
    await act(async () => {
      clock.advance(0);
      await new Promise<void>((resolve) => setImmediate(resolve));
    });
  }
  await flushPluginFallback(clock);
  expect(result.current.host).not.toBeNull();
  expect(opened).toBeDefined();
  const session = result.current.core.session!;
  expect(session.storyIds()).toEqual([]);
  expect(result.current.core.replicaReady).toBe(false);
  const replicaIndex = deferred.mock.calls.findIndex(([owner]) => owner === session);
  expect(replicaIndex).toBeGreaterThanOrEqual(0);
  const replica = deferred.mock.results[replicaIndex]!.value as ReturnType<typeof replicaHelpers.deferWorkerOpenReplica>;
  const snapshot = observePluginFallback(opened.documentRead({
    kind: 'readParagraphs', request: { view: 'accepted' },
  }), clock);
  await flushPluginFallback(clock);
  expect(snapshot.settled).toBe(true);
  expect(snapshot.error).toBeUndefined();
  if (!snapshot.value?.value.ok) throw new Error('The worker paragraph read did not succeed');
  const version = snapshot.value.version;
  const paragraphs = snapshot.value.value.paragraphs;
  expect(paragraphs.map((paragraph) => paragraph.text)).toEqual(['First paragraph', 'Tail paragraph']);
  const editor = result.current.pagedEditorRef.current!;
  const focus = mock(() => {});
  Object.assign(editor, { focus });
  const flush = spyOn(editor, 'flushPendingInput').mockImplementation(async () => {
    await replicaHelpers.awaitWorkerOpenReplica(session);
  });
  const start = spyOn(replica, 'start');
  const ensure = spyOn(replica, 'ensure');
  const requestReady = spyOn(replica, 'requestReady');
  const requested = spyOn(replicaHelpers, 'requestWorkerOpenReplica');
  const ensured = spyOn(replicaHelpers, 'ensureWorkerOpenReplica');
  const readinessRequested = spyOn(replicaHelpers, 'requestWorkerOpenReplicaReadiness');
  const ready = spyOn(replicaHelpers, 'awaitWorkerOpenReplica');
  const load = spyOn(session, 'loadState');
  const apply = spyOn<YrsSession, 'applyEdits'>(session, 'applyEdits');
  const selection = spyOn(session, 'setSelection');
  const binding = testBinding();
  binding.state.admission = async () => {
    const flushed = await flushEditorInput(result.current.pagedEditorRef);
    if (!flushed.ok) throw new DocxCommandAdmissionError(flushed.code);
  };
  const commands = createDocxCommandController();
  commands.attach(binding.binding);
  const controller = new AbortController();
  const lifetime = new AbortController();
  const invocation: PluginInvocation<DocxPluginSnapshot> = {
    pluginId: 'acme.review',
    activation: {},
    snapshot: {} as DocxPluginSnapshot,
    signal: controller.signal,
    lifetimeSignal: lifetime.signal,
    state: () => null,
    setState: () => false,
    onCleanup: () => {},
    run: async () => {},
    commit: (write) => write(),
    refusal: () => controller.signal.aborted ? 'aborted' : null,
  };
  const queries = {
    sourceState: () => ({ status: 'ready' }),
    anchorRect: () => ({ pageIndex: 0, x: 0, y: 0, width: 1, height: 1 }),
  } as unknown as DisplayListQueries;
  stampSourceVersion(queries, version);
  const clients = createPluginClients(invocation, {
    pagedEditorRef: result.current.pagedEditorRef,
    writeMode: () => 'editing',
    viewer: () => false,
    commands: () => commands,
    layout: () => ({ queries, complete: false, failed: false }),
    subscribeLayout: (listener) => session.onUpdate(listener),
  }, () => ({ document: 'write', editBatches: true }), commands.store);
  return {
    worker, frames, clock, tasks, visibility, result, unmount, session, replica,
    version, paragraphs, clients, focus, flush, start, ensure, requestReady,
    requested, ensured, readinessRequested, ready, load, apply, selection, binding,
  };
}

async function expectPluginOwnerFallback(
  calls: (env: Awaited<ReturnType<typeof openingPluginFallback>>) => readonly Promise<unknown>[],
  complete: (values: readonly unknown[], env: Awaited<ReturnType<typeof openingPluginFallback>>) => void | Promise<void>
) {
  const env = await openingPluginFallback();
  const { clock, replica, result, session, worker } = env;
  let outcomes: ReturnType<typeof observePluginFallback>[] = [];
  try {
    expect(clock.now).toBe(0);
    expect(worker.posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(result.current.pagedEditorRef.current?.hasPendingInput()).toBe(false);
    outcomes = calls(env).map((call) => observePluginFallback(call, clock));
    await flushPluginFallback(clock);
    expectUnstarted();
    await act(async () => { clock.advance(9_999); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(9_999);
    expectUnstarted();
    await act(async () => { clock.advance(2); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(10_001);
    expect(replica.started).toBe(true);
    expect(env.start).toHaveBeenCalledTimes(1);
    expect(result.current.renderer.layoutCompleteSession).toBeNull();
    expect(env.requested.mock.calls).toEqual([[session]]);
    expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(env.load).not.toHaveBeenCalled();
    expect(result.current.mainOpens).toEqual([]);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    expectPassiveCall();
    await act(async () => { worker.workers[0]!.release(); });
    await flushPluginFallback(clock);
    expect(env.tasks.tasks).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([false]);
    await act(async () => { await env.tasks.run(); });
    await flushPluginFallback(clock);
    expect(env.load).toHaveBeenCalledTimes(1);
    expect(env.tasks.tasks).toHaveLength(1);
    await act(async () => { await env.tasks.run(); });
    await flushPluginFallback(clock);
    expect(replica.hydrated).toBe(true);
    expect(replica.pending).toBe(true);
    const hydratedVersion = replica.readyVersion;
    expect(hydratedVersion).toBeDefined();
    if (hydratedVersion === undefined || replica.loadedVersion === undefined) {
      throw new Error('The replica did not record its hydrated versions');
    }
    expect(session.version()).toBe(hydratedVersion);
    expect(hydratedVersion).toBe(replica.loadedVersion);
    expect(hydratedVersion).not.toBe(env.version);
    expect(clock.now).toBe(10_001);
    expect(result.current.core.replicaReady).toBe(false);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    await act(async () => { clock.advance(2_999); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(13_000);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    await act(async () => { clock.advance(27_001); });
    await flushPluginFallback(clock);
    expect(clock.now).toBe(10_000 + 30_000 + 1);
    for (const outcome of outcomes) {
      expect(outcome.settled).toBe(true);
      expect(outcome.error).toBeUndefined();
      expect(outcome.at).toBeLessThanOrEqual(10_000 + 30_000 + 1);
    }
    expect(replica.pending).toBe(false);
    expect(result.current.core.replicaReady).toBe(true);
    const applied = env.apply.mock.results
      .flatMap((call) => call.type === 'return' && call.value ? [call.value] : [])
      .find((value) => value.ok && value.applied);
    expect(session.version()).toBe(applied?.version ?? hydratedVersion);
    expect(env.start).toHaveBeenCalledTimes(1);
    expect(env.requested.mock.calls).toEqual([[session]]);
    expect(env.load).toHaveBeenCalledTimes(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.core.session).toBe(session);
    expect<unknown[]>([null, session]).toContain(result.current.renderer.layoutCompleteSession);
    expect(result.current.errors).toEqual([]);
    expectPassiveCall();
    await complete(outcomes.map((outcome) => outcome.value), env);
  } finally {
    env.unmount();
    env.visibility.restore();
    env.tasks.restore();
    env.clock.restore();
    env.frames.restore();
  }

  function expectPassiveCall() {
    expect(env.ensure).not.toHaveBeenCalled();
    expect(env.requestReady).not.toHaveBeenCalled();
    expect(env.ensured).not.toHaveBeenCalled();
    expect(env.readinessRequested).not.toHaveBeenCalled();
    expect(env.ready.mock.calls).toEqual(outcomes.map(() => [session]));
    expect(env.flush).toHaveBeenCalledTimes(outcomes.length);
    expect(result.current.pagedEditorRef.current?.hasPendingInput()).toBe(false);
  }

  function expectUnstarted() {
    expect(replica.pending).toBe(true);
    expect(replica.started).toBe(false);
    expect(env.start).not.toHaveBeenCalled();
    expect(env.requested).not.toHaveBeenCalled();
    expect(env.load).not.toHaveBeenCalled();
    expect(env.apply).not.toHaveBeenCalled();
    expect(env.selection).not.toHaveBeenCalled();
    expect(env.focus).not.toHaveBeenCalled();
    expect(env.binding.calls).toEqual([]);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.renderer.layoutCompleteSession).toBeNull();
    expect(worker.posted.filter((request) => request.type === 'encodeState')).toEqual([]);
    for (const outcome of outcomes) expect(outcome.settled).toBe(false);
    expectPassiveCall();
  }
}

test('reads and versions issued after worker open wait for the owner fallback and complete by 40,001 ms without starting the peer', async () => {
  await expectPluginOwnerFallback((env) => [
    env.clients.read.version(),
    env.clients.read.readParagraphs({ view: 'accepted' }),
  ], (values, env) => {
    expect(values[0]).toEqual({ ok: true, version: env.replica.readyVersion });
    expect(values[1]).toEqual(env.session.readParagraphs({ view: 'accepted' }));
    expect(values[1]).toMatchObject({ ok: true, version: env.replica.readyVersion, paragraphs: [
      expect.objectContaining({ text: 'First paragraph' }),
      expect.objectContaining({ text: 'Tail paragraph' }),
    ] });
  });
});

test('text search issued after worker open waits for the owner fallback and completes by 40,001 ms without starting the peer', async () => {
  await expectPluginOwnerFallback((env) => [env.clients.read.findText({
    text: 'Tail', within: { kind: 'story', story: 'body' }, view: 'accepted',
  })], (values, env) => {
    expect(values[0]).toEqual(env.session.findText({
      text: 'Tail', within: { kind: 'story', story: 'body' }, view: 'accepted',
    }));
    expect(values[0]).toMatchObject({
      ok: true, version: env.replica.readyVersion,
      matches: [expect.objectContaining({ text: 'Tail' })],
    });
  });
});

test('navigation with focus issued after worker open waits for the owner fallback and completes by 40,001 ms without starting the peer', async () => {
  await expectPluginOwnerFallback((env) => [env.clients.navigation.scrollToParagraph(
    { story: 'body', paraId: env.paragraphs[1]!.paraId },
    { expectVersion: env.version, focus: true }
  )], (values, env) => {
    expect(values).toEqual([{ ok: true }]);
    expect(env.session.selection()?.head).toMatchObject({
      story: 'body', paraId: env.paragraphs[1]!.paraId, offset: 0,
    });
    expect(env.selection).toHaveBeenCalledTimes(1);
    expect(env.focus).toHaveBeenCalledTimes(1);
    expect(env.result.current.searchReveals).toHaveLength(1);
  });
});

test('mutations and commands issued after worker open wait for the owner fallback and complete by 40,001 ms without starting the peer', async () => {
  await expectPluginOwnerFallback((env) => [
    env.clients.edits!.applyEdits({
      expectVersion: env.version,
      steps: [{
        op: 'replaceText', target: { kind: 'paragraph', story: 'body', paraId: env.paragraphs[0]!.paraId },
        text: 'Owner-ready edit',
      }],
    }),
    env.clients.commands.execute('reviewNext', null),
  ], (values, env) => {
    const readyVersion = env.replica.readyVersion;
    if (readyVersion === undefined) throw new Error('The replica did not record its ready version');
    expect(values[0]).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
    expect(values[1]).toEqual({ ok: true, status: 'executed' });
    expect(env.session.paragraphs('body')[0]!.text).toBe('Owner-ready edit');
    expect(env.apply).toHaveBeenCalledTimes(1);
    expect(env.binding.calls).toEqual([{ id: 'reviewNext', args: null, ordered: true }]);
    expect(env.apply.mock.calls[0]![0].expectVersion).toBe(readyVersion);
  });
});

test('a pre-hydration worker version applies after unchanged owner hydration and becomes stale after an edit', async () => {
  await expectPluginOwnerFallback(() => [], async (_values, env) => {
    const readyVersion = env.replica.readyVersion;
    if (readyVersion === undefined) throw new Error('The replica did not record its ready version');
    const request = {
      expectVersion: env.version,
      steps: [{
        op: 'replaceText' as const,
        target: { kind: 'paragraph' as const, story: 'body', paraId: env.paragraphs[0]!.paraId },
        text: 'Owner-ready edit',
      }],
    };
    expect(env.replica.handoverVersion).toBe(env.version);
    expect(env.session.version()).toBe(readyVersion);
    const applied = await env.clients.edits!.applyEdits(request);
    expect(applied).toMatchObject({ ok: true, applied: true, changedStories: ['body'] });
    expect(env.apply.mock.calls[0]![0].expectVersion).toBe(readyVersion);
    expect(env.session.paragraphs('body')[0]!.text).toBe('Owner-ready edit');
    const editedVersion = env.session.version();
    expect(editedVersion).not.toBe(env.replica.readyVersion);
    expect(await env.clients.edits!.applyEdits({
      ...request, steps: [{ ...request.steps[0]!, text: 'Stale edit' }],
    })).toMatchObject({ ok: false, version: editedVersion, failure: { code: 'stale-version' } });
    expect(env.apply.mock.calls[1]![0].expectVersion).toBe(env.version);
    expect(env.session.version()).toBe(editedVersion);
    expect(env.session.paragraphs('body')[0]!.text).toBe('Owner-ready edit');
    expect(await env.clients.navigation.scrollToParagraph(
      { story: 'body', paraId: env.paragraphs[1]!.paraId },
      { expectVersion: env.version, focus: true }
    )).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(env.selection).not.toHaveBeenCalled();
    expect(env.focus).not.toHaveBeenCalled();
    expect(env.result.current.searchReveals).toEqual([]);
  });
});

test.each([false, true])('the ten-second fallback starts the editor peer when layout never settles with ownFrame=%s', async (ownFrame) => {
  const { posted, workers } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const fallback = holdPeerFallback();
  const visibility = stubDocumentVisibility('visible');
  const layoutReady = new Promise<void>(() => {});
  try {
    const { result } = renderHook(useHarness, {
      initialProps: { ...initialProps, layoutReady },
    });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    const load = spyOn(session, 'loadState');
    registerRestore(() => load.mockRestore());
    if (ownFrame) {
      act(() => result.current.pipeline.runLayoutPipeline());
      await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
      act(() => result.current.presentFrame());
      await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
    }
    act(() => frames.runIdle());
    act(() => fallback.advance(9999));
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    expect(load).not.toHaveBeenCalled();
    act(() => fallback.advance(1));
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(true);
    await waitFor(() => expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(fallback.timers.size).toBe(0);
    act(() => frames.runIdle());
    act(() => fallback.advance(10_000));
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    await act(async () => {
      workers[0].release();
      await awaitWorkerOpenReplica(session);
    });
    expect(result.current.core.replicaReady).toBe(true);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      visibility.restore();
      fallback.restore();
      frames.restore();
    }
  }
});

test.each(['layout', 'idle'] as const)('replacing a document cancels the pending peer start during %s', async (stage) => {
  const { posted } = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const fallback = holdPeerFallback();
  const visibility = stubDocumentVisibility('visible');
  let releaseLayout!: () => void;
  const layoutReady = new Promise<void>((resolve) => { releaseLayout = resolve; });
  try {
    const { result, rerender } = renderHook(useHarness, {
      initialProps: { ...initialProps, layoutReady: stage === 'layout' ? layoutReady : undefined },
    });
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const previous = result.current.core.session!;
    const load = spyOn(previous, 'loadState');
    registerRestore(() => load.mockRestore());
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(previous));
    act(() => result.current.presentFrame());
    await frames.untilCommitted(result.current.renderer.settledDisplayList(null, null, 'window'));
    expect(replicaHelpers.workerOpenReplicaStarted(previous)).toBe(false);
    const idleCallbacks = [...frames.idleCallbacks.values()].filter(({ options }) => options?.timeout === 2000);
    expect(idleCallbacks).toHaveLength(stage === 'idle' ? 1 : 0);
    const staleFallbacks = [...fallback.timers.values()];
    expect(staleFallbacks).toHaveLength(1);
    act(() => rerender({ ...initialProps, source: bytes.slice(), generation: 2, layoutReady: undefined }));
    await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
    expect(result.current.core.session).not.toBe(previous);
    expect([...frames.idleCallbacks.values()].filter(({ options }) => options?.timeout === 2000)).toHaveLength(0);
    await act(async () => {
      releaseLayout();
      for (const { callback } of idleCallbacks) callback({ didTimeout: false, timeRemaining: () => 50 });
      for (const { callback } of staleFallbacks) callback();
      frames.runIdle();
    });
    expect(replicaHelpers.workerOpenReplicaPending(previous)).toBe(false);
    expect(load).not.toHaveBeenCalled();
    expect(posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([]);
    expect(result.current.errors).toEqual([]);
  } finally {
    try {
      cleanup();
    } finally {
      visibility.restore();
      fallback.restore();
      frames.restore();
    }
  }
});

test('completion slices do not register an editor waiter or prefetch before the layout-complete signal', async () => {
  const options = { holdCompletion: true };
  const worker = installWorker(options);
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const onLayoutWait = mock(() => {});
  const props: HarnessProps = {
    ...initialProps, source: await longFixture(1200), followCompletion: true,
    layoutCompleteSession: null, onLayoutWait,
  };
  const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const session = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.presentedEngine).toBe(session));
    await worker.sent('completeLayout');
    act(() => result.current.presentFrame());
    await act(async () => {
      frames.run();
      frames.runIdle();
    });
    expect(result.current.renderer.pendingCompletion).toBeNull();
    expect(result.current.renderer.layoutCompleteSession).toBeNull();
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    options.holdCompletion = false;
    await act(async () => worker.workers[0].release());
    await waitFor(() => expect(result.current.renderer.layoutCompleteSession).toBe(session));
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => rerender({ ...props, layoutCompleteSession: session }));
    await frames.waitFor(() => expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1));
    expect(onLayoutWait).toHaveBeenCalledTimes(1);
    expectStatePrefetchAfterPageBuilds(worker.posted);
    expect(replicaHelpers.workerOpenReplicaStarted(session)).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    await frames.waitFor(() => expect([...frames.idleCallbacks.values()]
      .filter(({ options }) => options?.timeout === 2000)).toHaveLength(1));
    await act(async () => frames.runIdle());
    await waitFor(() => expect(result.current.core.replicaReady).toBe(true));
    expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(result.current.mainOpens).toEqual([false]);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    visibility.restore();
    frames.restore();
  }
}, 15_000);

test('a keystroke during completion slices starts the peer and replays once after readiness', async () => {
  const opened = await openingEditor(true, false, { holdFullLayout: true });
  const tasks = holdHydrationTasks();
  let insert: ReturnType<typeof spyOn<YrsSession, 'insertText'>> | undefined;
  try {
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    expect(opened.harness.renderer.layoutCompleteSession).toBeNull();
    expect(replicaHelpers.workerOpenReplicaStarted(full)).toBe(false);
    insert = spyOn(full, 'insertText');
    opened.click(6);
    opened.type('K');
    expect(replicaHelpers.workerOpenReplicaStarted(full)).toBe(true);
    expect(insert).not.toHaveBeenCalled();
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    await finishHeldHydration(opened, tasks);
    expect(replicaHelpers.workerOpenReplicaHydrating(full)).toBe(false);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(insert).not.toHaveBeenCalled();
    const completion = await opened.received('completeLayout');
    await act(async () => {
      opened.reply(completion);
      await awaitWorkerOpenReplica(full);
      await opened.editor.current!.flushPendingInput();
    });
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(insert).toHaveBeenCalledTimes(1);
    expect(full.paragraphs('body')[0].text).toBe('FirstK paragraph');
    await act(async () => opened.editor.current!.flushPendingInput());
    expect(insert).toHaveBeenCalledTimes(1);
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    insert?.mockRestore();
    opened.close();
    tasks.restore();
  }
});

test.each(['ctrlKey', 'metaKey'] as const)('keyboard and menu select-all during completion slices wait for layout-complete with %s', async (modifier) => {
  const opened = await openingEditor(true, false, { holdFullLayout: true });
  const tasks = holdHydrationTasks();
  try {
    act(() => opened.editor.current!.selectAll());
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    const full = await opened.switchToFull();
    await opened.presentFull(full, false);
    expect(opened.harness.renderer.layoutCompleteSession).toBeNull();
    expect(replicaHelpers.workerOpenReplicaStarted(full)).toBe(false);
    act(() => opened.editor.current!.selectAll());
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
    });
    expect(replicaHelpers.workerOpenReplicaStarted(full)).toBe(false);
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    const textarea = opened.view.getByTestId('yrs-input');
    act(() => textarea.focus());
    fireEvent.keyDown(textarea, { key: 'a', [modifier]: true });
    expect(replicaHelpers.workerOpenReplicaStarted(full)).toBe(false);
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    const completion = await opened.received('completeLayout');
    await act(async () => opened.reply(completion));
    await opened.frames.waitFor(() => expect(opened.harness.renderer.layoutCompleteSession).toBe(full));
    await opened.frames.waitFor(() => expect([...opened.frames.idleCallbacks.values()]
      .filter(({ options }) => options?.timeout === 2000)).toHaveLength(1));
    await act(async () => opened.frames.runIdle());
    await finishHeldHydration(opened, tasks);
    await act(async () => {
      await awaitWorkerOpenReplica(full);
      await opened.editor.current!.flushPendingInput();
    });
    const paragraphs = full.paragraphs('body');
    expect(full.selection()).toEqual({
      anchor: { story: 'body', paraId: paragraphs[0].paraId, offset: 0 },
      head: { story: 'body', paraId: paragraphs.at(-1)!.paraId, offset: paragraphs.at(-1)!.text.length },
    });
    expect(document.activeElement).toBe(textarea);
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    tasks.restore();
  }
});

function openingClipboard(write: (text: string) => Promise<void> = async () => {}) {
  const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const item = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
  const writeText = mock(write);
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
  Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: undefined });
  const restore = registerRestore(() => {
    if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
    else Reflect.deleteProperty(navigator, 'clipboard');
    if (item) Object.defineProperty(globalThis, 'ClipboardItem', item);
    else Reflect.deleteProperty(globalThis, 'ClipboardItem');
  });
  return { writeText, restore };
}

async function editorWithoutLayoutCompleteSignal(holdInput = false, resolveDisplayListQueries?: ResolveDisplayListQueries) {
  const worker = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const onLayoutWait = mock(() => {});
  const source = await longFixture(1);
  const editor = createRef<PagedEditorRef>();
  const canvasHost = createRef<HTMLDivElement>();
  let harness!: ReturnType<typeof useHarness>;
  function Editable({ held, layoutCompleteSession = null }: { held: boolean; layoutCompleteSession?: YrsSession | null }) {
    harness = useHarness({ ...initialProps, source, layoutCompleteSession, onLayoutWait });
    return <>
      <div ref={canvasHost} className="canvas-pages"><canvas className="canvas-page" data-page-index="0" /></div>
      <PagedEditor ref={editor} document={harness.host?.document ?? null} yrsCore={harness.core}
        readOnly={held} holdInput={held} inputScope={1}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(font.buffer as ArrayBuffer) }}
        fontRequirementsInWorker={harness.renderer.fontRequirementsInWorker}
        layoutInWorker={harness.renderer.layoutInWorker}
        canvasHostRef={canvasHost} displayListQueries={harness.renderer.queries}
        inputQueries={harness.renderer.inputQueries} resolveDisplayListQueries={resolveDisplayListQueries} />
    </>;
  }
  if (!document.fonts) Object.defineProperty(document, 'fonts', {
    value: { addEventListener: () => {}, removeEventListener: () => {} }, configurable: true,
  });
  const view = render(<Editable held={holdInput} />);
  const close = () => {
    view.unmount();
    visibility.restore();
    frames.restore();
  };
  try {
    await waitFor(() => expect(harness.host).not.toBeNull());
    const session = harness.core.session!;
    act(() => harness.pipeline.runLayoutPipeline());
    await waitFor(() => expect(harness.renderer.presentedEngine).toBe(session));
    await waitFor(() => expect(harness.renderer.queries?.isReady()).toBe(true));
    act(() => harness.presentFrame());
    const canvas = canvasHost.current!.firstElementChild!;
    const size = harness.renderer.queries!.pageSize(0)!;
    canvas.getBoundingClientRect = () => ({ left: 0, top: 0, right: size.width,
      bottom: size.height, ...size }) as DOMRect;
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(false);
    return {
      ...worker, frames, view, session, editor, canvas, onLayoutWait, close,
      get harness() { return harness; },
      releaseHeldInput(layoutCompleteSession: YrsSession | null = null) {
        view.rerender(<Editable held={false} layoutCompleteSession={layoutCompleteSession} />);
      },
    };
  } catch (error) {
    close();
    throw error;
  }
}

function selectOpeningText(opened: Awaited<ReturnType<typeof editorWithoutLayoutCompleteSignal>>) {
  const point = (position: number) => {
    const caret = opened.harness.renderer.queries!.caretRect(position)!;
    return { clientX: caret.x, clientY: caret.y + caret.height / 2, button: 0, detail: 1 };
  };
  fireEvent.mouseDown(opened.canvas, point(1));
  fireEvent.mouseUp(window, point(6));
}

async function readyEditorWithPendingNavigation() {
  let releaseNavigation!: () => void;
  const blocked = new Promise<null>((resolve) => { releaseNavigation = () => resolve(null); });
  let started!: () => void;
  const resolving = new Promise<void>((resolve) => { started = resolve; });
  const resolveQueries = mock(() => {
    started();
    return blocked;
  });
  const opened = await editorWithoutLayoutCompleteSignal(false, resolveQueries);
  try {
    act(() => opened.releaseHeldInput(opened.session));
    await opened.frames.waitFor(() => expect([...opened.frames.idleCallbacks.values()]
      .filter(({ options }) => options?.timeout === 2000)).toHaveLength(1));
    await act(async () => opened.frames.runIdle());
    await opened.sent('encodeState');
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
      await opened.editor.current!.flushPendingInput();
    });
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    return {
      opened, releaseNavigation,
      pointAt(position: number) {
        const caret = opened.harness.renderer.queries!.caretRect(position)!;
        return { clientX: caret.x, clientY: caret.y + caret.height / 2, button: 0, detail: 1 };
      },
      async startNavigation() {
        await act(async () => {
          fireEvent.keyDown(opened.view.getByTestId('yrs-input'), { key: 'ArrowDown' });
          await resolving;
        });
        expect(resolveQueries).toHaveBeenCalledTimes(1);
        expect(opened.harness.core.replicaReady).toBe(true);
        expect(opened.harness.renderer.inputQueries!.isReady()).toBe(true);
        expect(opened.editor.current!.hasPendingInput()).toBe(true);
      },
    };
  } catch (error) {
    releaseNavigation();
    opened.close();
    throw error;
  }
}

test.each([false, true])('a queued keystroke starts the peer before the layout signal and replays once with heldInput=%s', async (held) => {
  const opened = await editorWithoutLayoutCompleteSignal(held);
  const insert = spyOn(opened.session, 'insertText');
  const load = spyOn(opened.session, 'loadState');
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    fireEvent.keyDown(textarea, { key: 'A' });
    fireEvent.input(textarea, { target: { value: 'A' } });
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(true);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(insert).not.toHaveBeenCalled();
    expect(load).not.toHaveBeenCalled();
    expect(opened.onLayoutWait).not.toHaveBeenCalled();
    await opened.sent('encodeState');
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    act(() => opened.releaseHeldInput());
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
      await opened.editor.current!.flushPendingInput();
    });
    expect(load).toHaveBeenCalledTimes(1);
    expect(insert).toHaveBeenCalledTimes(1);
    expect(opened.session.paragraphs('body')[0].text).toBe('AFirst paragraph');
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
      await opened.editor.current!.flushPendingInput();
    });
    expect(insert).toHaveBeenCalledTimes(1);
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    insert.mockRestore();
    load.mockRestore();
    opened.close();
  }
});

test.each(['paste', 'cut', 'composition'] as const)('queued %s starts the peer immediately without a layout waiter', async (kind) => {
  const opened = await editorWithoutLayoutCompleteSignal();
  const insert = spyOn(opened.session, 'insertText');
  const remove = spyOn(opened.session, 'deleteRange');
  const clipboard = openingClipboard();
  try {
    const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
    act(() => textarea.focus());
    if (kind === 'cut') selectOpeningText(opened);
    if (kind === 'paste') fireEvent.paste(textarea, { clipboardData: { getData: () => 'P' } });
    else if (kind === 'cut') fireEvent.cut(textarea);
    else fireEvent.compositionStart(textarea);
    if (kind === 'composition') {
      textarea.value = 'I';
      fireEvent.compositionEnd(textarea, { data: 'I' });
    }
    expect(document.activeElement).toBe(textarea);
    expect(insert).not.toHaveBeenCalled();
    expect(remove).not.toHaveBeenCalled();
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(true);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.onLayoutWait).not.toHaveBeenCalled();
    await opened.sent('encodeState');
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.mainOpens).toEqual([]);
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
      await opened.editor.current!.flushPendingInput();
    });
    const assertReplay = () => {
      expect(opened.harness.core.replicaReady).toBe(true);
      expect(insert).toHaveBeenCalledTimes(kind === 'paste' || kind === 'composition' ? 1 : 0);
      expect(remove).toHaveBeenCalledTimes(kind === 'cut' ? 1 : 0);
      expect(opened.session.paragraphs('body')[0].text).toBe(
        kind === 'paste' ? 'PFirst paragraph' : kind === 'composition' ? 'IFirst paragraph' :
          ' paragraph'
      );
      if (kind === 'cut') {
        expect(clipboard.writeText).toHaveBeenCalledTimes(1);
        expect(clipboard.writeText).toHaveBeenCalledWith('First');
      }
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
      expect(document.activeElement).toBe(textarea);
    };
    assertReplay();
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
      await opened.editor.current!.flushPendingInput();
    });
    assertReplay();
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    insert.mockRestore();
    remove.mockRestore();
    opened.close();
    clipboard.restore();
  }
});

test.each([false, true].flatMap((held) =>
  ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'PageUp', 'PageDown', 'Home', 'End'].flatMap((key) =>
    [false, true].map((shift) => [key, shift, held] as const)
  )
))('queued %s does not start the peer before layout-complete with Shift=%s and heldInput=%s', async (key, shift, held) => {
  const opened = await editorWithoutLayoutCompleteSignal(held);
  const insert = spyOn(opened.session, 'insertText');
  const remove = spyOn(opened.session, 'deleteRange');
  const select = spyOn(opened.session, 'setSelection');
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    act(() => textarea.focus());
    fireEvent.keyDown(textarea, { key, shiftKey: shift });
    expect(document.activeElement).toBe(textarea);
    expect(insert).not.toHaveBeenCalled();
    expect(remove).not.toHaveBeenCalled();
    act(() => opened.releaseHeldInput());
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
    });
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(false);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.onLayoutWait).not.toHaveBeenCalled();
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(opened.harness.mainOpens).toEqual([]);
    act(() => opened.releaseHeldInput(opened.session));
    await opened.frames.waitFor(() => expect([...opened.frames.idleCallbacks.values()]
      .filter(({ options }) => options?.timeout === 2000)).toHaveLength(1));
    await act(async () => opened.frames.runIdle());
    await opened.sent('encodeState');
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
      await opened.editor.current!.flushPendingInput();
    });
    const assertReplay = () => {
      expect(opened.harness.core.replicaReady).toBe(true);
      expect(insert).toHaveBeenCalledTimes(0);
      expect(remove).toHaveBeenCalledTimes(0);
      if (key === 'ArrowRight') {
        expect(select.mock.calls.filter(([anchor, head]) => (head ?? anchor).offset === 1)).toHaveLength(1);
        expect(opened.session.selection()?.head.offset).toBe(1);
      }
      expect(opened.session.paragraphs('body')[0].text).toBe('First paragraph');
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
      expect(document.activeElement).toBe(textarea);
    };
    assertReplay();
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
      await opened.editor.current!.flushPendingInput();
    });
    assertReplay();
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    insert.mockRestore();
    remove.mockRestore();
    select.mockRestore();
    opened.close();
  }
});

test.each([false, true])('ArrowRight and PageDown wait for a typed character to start the peer and queued input replays once in order with heldInput=%s', async (held) => {
  const opened = await editorWithoutLayoutCompleteSignal(held);
  const replay: string[] = [];
  const setSelection = opened.session.setSelection.bind(opened.session);
  const insertText = opened.session.insertText.bind(opened.session);
  const select = spyOn(opened.session, 'setSelection').mockImplementation((anchor, head) => {
    if ((head ?? anchor).offset === 1) replay.push('ArrowRight');
    return setSelection(anchor, head);
  });
  const insert = spyOn(opened.session, 'insertText').mockImplementation((...args) => {
    replay.push(args[1]);
    return insertText(...args);
  });
  const load = spyOn(opened.session, 'loadState');
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    act(() => textarea.focus());
    for (const key of ['ArrowRight', 'PageDown']) {
      fireEvent.keyDown(textarea, { key });
      expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(false);
      expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
      expect(document.activeElement).toBe(textarea);
    }
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    fireEvent.keyDown(textarea, { key: 'A' });
    fireEvent.input(textarea, { target: { value: 'A' } });
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(true);
    expect(opened.harness.core.replicaReady).toBe(false);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(replay).toEqual([]);
    expect(insert).not.toHaveBeenCalled();
    expect(load).not.toHaveBeenCalled();
    expect(opened.onLayoutWait).not.toHaveBeenCalled();
    expect(document.activeElement).toBe(textarea);
    await opened.sent('encodeState');
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.mainOpens).toEqual([]);
    act(() => opened.releaseHeldInput());
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
      await opened.editor.current!.flushPendingInput();
    });
    const assertReplay = () => {
      expect(opened.harness.core.replicaReady).toBe(true);
      expect(load).toHaveBeenCalledTimes(1);
      expect(insert).toHaveBeenCalledTimes(1);
      expect(select.mock.calls.filter(([anchor, head]) => (head ?? anchor).offset === 1)).toHaveLength(1);
      expect(replay).toEqual(['ArrowRight', 'A']);
      expect(opened.session.paragraphs('body')[0].text).toBe('FAirst paragraph');
      expect(opened.session.selection()?.head.offset).toBe(2);
      expect(opened.editor.current!.hasPendingInput()).toBe(false);
      expect(document.activeElement).toBe(textarea);
    };
    assertReplay();
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
      await opened.editor.current!.flushPendingInput();
    });
    assertReplay();
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    select.mockRestore();
    insert.mockRestore();
    load.mockRestore();
    opened.close();
  }
});

test.each([false, true])('a queued cut writes the selected text before deletion and replays once with heldInput=%s', async (held) => {
  const opened = await editorWithoutLayoutCompleteSignal(held);
  let releaseWrite!: () => void;
  const writeReady = new Promise<void>((resolve) => { releaseWrite = resolve; });
  let copied = '';
  const clipboard = openingClipboard(async (text) => {
    copied = text;
    await writeReady;
  });
  const remove = spyOn(opened.session, 'deleteRange');
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    act(() => textarea.focus());
    selectOpeningText(opened);
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(false);
    fireEvent.cut(textarea);
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(true);
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    expect(clipboard.writeText).not.toHaveBeenCalled();
    expect(remove).not.toHaveBeenCalled();
    await opened.sent('encodeState');
    act(() => opened.releaseHeldInput());
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
    });
    await waitFor(() => expect(clipboard.writeText).toHaveBeenCalledWith('First'));
    expect(clipboard.writeText).toHaveBeenCalledTimes(1);
    expect(copied).toBe('First');
    expect(opened.session.paragraphs('body')[0].text).toBe('First paragraph');
    expect(remove).not.toHaveBeenCalled();
    await act(async () => {
      releaseWrite();
      await opened.editor.current!.flushPendingInput();
    });
    expect(remove).toHaveBeenCalledTimes(1);
    expect(opened.session.paragraphs('body')[0].text).toBe(' paragraph');
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(document.activeElement).toBe(textarea);
    await act(async () => opened.editor.current!.flushPendingInput());
    expect(remove).toHaveBeenCalledTimes(1);
    expect(clipboard.writeText).toHaveBeenCalledTimes(1);
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    releaseWrite();
    remove.mockRestore();
    opened.close();
    clipboard.restore();
  }
});

test('a held copy writes the held selection made before it', async () => {
  const opened = await editorWithoutLayoutCompleteSignal(true);
  const clipboard = openingClipboard();
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    act(() => textarea.focus());
    selectOpeningText(opened);
    fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true });
    expect(clipboard.writeText).not.toHaveBeenCalled();
    act(() => opened.releaseHeldInput());
    act(() => opened.releaseHeldInput(opened.session));
    await opened.frames.waitFor(() => expect([...opened.frames.idleCallbacks.values()]
      .filter(({ options }) => options?.timeout === 2000)).toHaveLength(1));
    await act(async () => opened.frames.runIdle());
    await opened.sent('encodeState');
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
      await opened.editor.current!.flushPendingInput();
    });
    await waitFor(() => expect(clipboard.writeText).toHaveBeenCalledWith('First'));
    expect(clipboard.writeText).toHaveBeenCalledTimes(1);
    expect(opened.session.paragraphs('body')[0].text).toBe('First paragraph');
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
    clipboard.restore();
  }
});

test('a ready copy queued behind navigation writes the new selection once and an idle copy stays synchronous', async () => {
  const { opened, pointAt, startNavigation, releaseNavigation } = await readyEditorWithPendingNavigation();
  const clipboard = openingClipboard();
  try {
    const textarea = opened.view.getByTestId('yrs-input') as HTMLTextAreaElement;
    act(() => textarea.focus());
    fireEvent.mouseDown(opened.canvas, pointAt(1));
    fireEvent.mouseMove(window, pointAt(6));
    fireEvent.mouseUp(window, pointAt(6));
    await act(async () => opened.editor.current!.flushPendingInput());
    expect(opened.session.selection()?.anchor.offset).toBe(0);
    expect(opened.session.selection()?.head.offset).toBe(5);
    await startNavigation();
    const point = { ...pointAt(9), detail: 2 };
    fireEvent.mouseDown(opened.canvas, point);
    fireEvent.mouseUp(window, point);
    fireEvent.click(opened.canvas, point);
    fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true });
    expect(opened.session.selection()?.anchor.offset).toBe(0);
    expect(opened.session.selection()?.head.offset).toBe(5);
    expect(clipboard.writeText).not.toHaveBeenCalled();
    expect(opened.editor.current!.hasPendingInput()).toBe(true);
    await act(async () => {
      releaseNavigation();
      await opened.editor.current!.flushPendingInput();
    });
    await waitFor(() => expect(clipboard.writeText.mock.calls).toEqual([['paragraph']]));
    expect(opened.session.selection()?.anchor.offset).toBe(6);
    expect(opened.session.selection()?.head.offset).toBe(15);
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
      await opened.editor.current!.flushPendingInput();
    });
    expect(clipboard.writeText.mock.calls).toEqual([['paragraph']]);
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    act(() => {
      expect(fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true })).toBe(true);
      expect(textarea.value).toBe('paragraph');
      expect(textarea.selectionStart).toBe(0);
      expect(textarea.selectionEnd).toBe(9);
    });
    await act(async () => { await Promise.resolve(); });
    expect(clipboard.writeText.mock.calls).toEqual([['paragraph']]);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    releaseNavigation();
    opened.close();
    clipboard.restore();
  }
});

test.each([false, true])('a queued cut whose clipboard write fails keeps the selection with heldInput=%s', async (held) => {
  const opened = await editorWithoutLayoutCompleteSignal(held);
  const clipboard = openingClipboard(async () => {
    throw new Error('denied');
  });
  const remove = spyOn(opened.session, 'deleteRange');
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    act(() => textarea.focus());
    selectOpeningText(opened);
    fireEvent.cut(textarea);
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(true);
    await opened.sent('encodeState');
    act(() => opened.releaseHeldInput());
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
    });
    await waitFor(() => expect(clipboard.writeText).toHaveBeenCalledWith('First'));
    await act(async () => opened.editor.current!.flushPendingInput());
    expect(remove).not.toHaveBeenCalled();
    expect(opened.session.paragraphs('body')[0].text).toBe('First paragraph');
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    remove.mockRestore();
    opened.close();
    clipboard.restore();
  }
});

test.each([false, true])('a cut without the Clipboard API queues no deletion with heldInput=%s', async (held) => {
  const opened = await editorWithoutLayoutCompleteSignal(held);
  const descriptor = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
  const restore = registerRestore(() => {
    if (descriptor) Object.defineProperty(navigator, 'clipboard', descriptor);
    else Reflect.deleteProperty(navigator, 'clipboard');
  });
  const remove = spyOn(opened.session, 'deleteRange');
  try {
    const textarea = opened.view.getByTestId('yrs-input');
    act(() => textarea.focus());
    selectOpeningText(opened);
    fireEvent.cut(textarea);
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(false);
    act(() => opened.releaseHeldInput());
    act(() => opened.releaseHeldInput(opened.session));
    await opened.frames.waitFor(() => expect([...opened.frames.idleCallbacks.values()]
      .filter(({ options }) => options?.timeout === 2000)).toHaveLength(1));
    await act(async () => opened.frames.runIdle());
    await opened.sent('encodeState');
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
      await opened.editor.current!.flushPendingInput();
    });
    expect(remove).not.toHaveBeenCalled();
    expect(opened.session.paragraphs('body')[0].text).toBe('First paragraph');
    expect(opened.editor.current!.hasPendingInput()).toBe(false);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    remove.mockRestore();
    opened.close();
    restore();
  }
});

test.each([false, true])('pointer moves, mouse selection, wheel and scroll do not start the opening peer with heldInput=%s', async (held) => {
  const opened = await editorWithoutLayoutCompleteSignal(held);
  try {
    const caret = opened.harness.renderer.queries!.caretRect(6)!;
    const point = { clientX: caret.x, clientY: caret.y + caret.height / 2, button: 0 };
    fireEvent.pointerMove(opened.canvas, point);
    fireEvent.mouseDown(opened.canvas, point);
    fireEvent.mouseMove(opened.canvas, point);
    fireEvent.mouseUp(window, point);
    fireEvent.click(opened.canvas, point);
    fireEvent.wheel(opened.canvas, { deltaY: 100 });
    fireEvent.scroll(opened.canvas);
    fireEvent.keyDown(opened.view.getByTestId('yrs-input'), { key: 'Shift' });
    act(() => opened.releaseHeldInput());
    await act(async () => {
      opened.frames.run();
      opened.frames.runIdle();
    });
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(false);
    expect(opened.onLayoutWait).not.toHaveBeenCalled();
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(opened.harness.mainOpens).toEqual([]);
    expect(opened.harness.errors).toEqual([]);
  } finally {
    opened.close();
  }
});

test('the ten-second fallback starts the peer when a presented frame has no layout-complete signal', async () => {
  const fallback = holdPeerFallback();
  const opened = await editorWithoutLayoutCompleteSignal();
  try {
    act(() => fallback.advance(9999));
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(false);
    expect(opened.onLayoutWait).not.toHaveBeenCalled();
    expect(opened.posted.some((request) => request.type === 'encodeState')).toBe(false);
    act(() => fallback.advance(1));
    expect(replicaHelpers.workerOpenReplicaStarted(opened.session)).toBe(true);
    expect(opened.onLayoutWait).not.toHaveBeenCalled();
    await opened.sent('encodeState');
    expect(opened.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(fallback.timers.size).toBe(0);
    await act(async () => {
      opened.workers[0].release();
      await awaitWorkerOpenReplica(opened.session);
    });
    expect(opened.harness.core.replicaReady).toBe(true);
    expect(opened.harness.mainOpens).toEqual([false]);
  } finally {
    opened.close();
    fallback.restore();
  }
});

test('a stale layout-complete signal cannot start the replacement session peer', async () => {
  const worker = installWorker({ holdState: true });
  const frames = holdFrames(true);
  const visibility = stubDocumentVisibility('visible');
  const onLayoutWait = mock(() => {});
  const props: HarnessProps = { ...initialProps, source: await longFixture(1), layoutCompleteSession: null, onLayoutWait };
  const { result, rerender, unmount } = renderHook(useHarness, { initialProps: props });
  try {
    await waitFor(() => expect(result.current.host).not.toBeNull());
    const previous = result.current.core.session!;
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.layoutCompleteSession).toBe(previous));
    act(() => result.current.presentFrame());
    const replacementProps = { ...props, source: await longFixture(2), generation: 2, layoutCompleteSession: previous };
    act(() => rerender(replacementProps));
    await waitFor(() => expect(result.current.core.sessionGeneration).toBe(2));
    const replacement = result.current.core.session!;
    expect(replacement).not.toBe(previous);
    expect(result.current.renderer.layoutCompleteSession).not.toBe(previous);
    act(() => result.current.pipeline.runLayoutPipeline());
    await waitFor(() => expect(result.current.renderer.layoutCompleteSession).toBe(replacement));
    act(() => result.current.presentFrame());
    await act(async () => {
      frames.run();
      frames.runIdle();
    });
    expect(onLayoutWait).not.toHaveBeenCalled();
    expect(replicaHelpers.workerOpenReplicaStarted(replacement)).toBe(false);
    expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(result.current.mainOpens).toEqual([]);
    act(() => rerender({ ...replacementProps, layoutCompleteSession: replacement }));
    await frames.waitFor(() => expect(onLayoutWait).toHaveBeenCalledTimes(1));
    await worker.sent('encodeState');
    await frames.waitFor(() => expect([...frames.idleCallbacks.values()]
      .filter(({ options }) => options?.timeout === 2000)).toHaveLength(1));
    await act(async () => frames.runIdle());
    expect(replicaHelpers.workerOpenReplicaStarted(replacement)).toBe(true);
    expect(worker.posted.filter((request) => request.type === 'encodeState')).toHaveLength(1);
    expect(result.current.errors).toEqual([]);
  } finally {
    unmount();
    visibility.restore();
    frames.restore();
  }
});
