import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { createRef, startTransition, Suspense, useEffect, useLayoutEffect, useState } from 'react';
import * as measurement from '@betteroffice/docx/layout';
import * as utils from '@betteroffice/docx/utils';
import * as yrs from '@betteroffice/docx/yrs';
import * as peerless from '@betteroffice/docx/peerless';
import * as frameRender from '@betteroffice/docx/layout/render';
import type {
  DocxCommandFrame,
  DocxWorkerEditSession,
  ResidentCommandInitReply,
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs';
import docxPackage from '../../../../../docx/package.json';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef, type DocxErrorInfo } from '../../DocxEditor';
import { defineDocxPlugin } from '../../../plugins/defineDocxPlugin';
import type { DocxPluginContext } from '../../../plugins/types';
import * as core from '../hooks/useYrsCoreSession';
import * as display from '../hooks/useDisplayList';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { PeerlessEchoIntent } from './echoPainter';
import type { PeerlessFrameGeometry } from './viewStore';
import * as paged from './PeerlessPagedEditor';
import type { PeerlessViewSnapshot } from './viewStore';
import { OUTLINE_RESERVED_SPACE } from '../../DocumentOutline';
import { SIDEBAR_DOCUMENT_SHIFT } from '../../sidebar/constants';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from '../internals/engineChoice';
import { LEGACY_SYNC_API_PROP, PeerlessSyncApiError } from './contracts';

import { createInputAdmission } from './inputAdmission';
import { createPeerlessCommandBinding } from './commandBinding';
import { createEchoPainter } from './echoPainter';
import { createFramePresentation } from './framePresentation';
import { inputs, proposalHost } from './workerProposalHost.testing';
import { peerlessPreviewPaint } from './WorkerDocumentPreview';
import * as clipboard from './clipboard';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const source = Uint8Array.of(1, 2, 3);
const peerlessProps: DocxEditorProps = {
  documentBuffer: source,
};
const requests: ResidentEngineWorkerRequest[] = [];
const workers: FakeWorker[] = [];
let coreSession: ReturnType<typeof installCore>;
let surface: ReturnType<typeof installSurface>;
let binding: PeerlessEditSessionBinding | undefined;
let hostJson: string;
let fontRequirements: measurement.ResidentFontRequirement[];
let holdInitial = false;
let holdPreviewPaint = false;
let failPreview = false;
let refuseCut = false;
let holdCut = false;
let failWholeOpen = false;
/** Per source preview, in order: whether its first page is built or missing. */
let previewPages: ('built' | 'none')[] = [];
const PREVIEW_EPOCHS = { built: 90, none: 1 } as const;

function hostMetadata() {
  return {
    referencedFonts: [] as string[],
    envelope: {
      wireVersion: 1,
      document: {
        package: {
          document: { content: [], sections: [] as {
            id: string; contentStart: number; contentEnd: number; properties: Record<string, unknown>;
          }[] },
          theme: {}, numbering: {}, settings: {}, fontTable: { fonts: [] },
          relationshipEntries: [], mediaEntries: [], chartEntries: [],
        },
      },
      embeddedFontParts: [],
    },
  };
}

class FakeWorker {
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  terminate = mock(() => {});
  initial: ResidentCommandInitReply | undefined;
  text = '';
  holdEdits = false;
  refuseNextEdit = false;
  refused = new Set<number>();
  fontRevision = 0;
  pendingReplies: ResidentEngineWorkerResponse[] = [];

  constructor() {
    workers.push(this);
  }

  postMessage(message: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule) {
    if (!('id' in message)) return;
    requests.push(message);
    let response: ResidentEngineWorkerResponse;
    if (message.type === 'openSource' && message.source.cut && refuseCut) {
      response = { id: message.id, ok: true, previewRefused: true };
    } else if (message.type === 'openSource' && !message.source.cut && failWholeOpen) {
      response = { id: message.id, ok: false, error: 'Whole source failed to open' };
    } else if (message.type === 'openSource') {
      this.text = '';
      response = { id: message.id, ok: true, hostJson, sourceDocument: yrs.decodeDocxHostJson(hostJson, source).document, stateVector: new ArrayBuffer(0),
        ...(message.source.preview ? { frame: commandFrame(1).payload.frame, documentPreview: true,
          caret: { frameEpoch: 1, caretRect: null }, selection: null, layoutRevision: 1 } : {}) };
    } else if (message.type === 'previewSource') {
      response = failPreview ? { id: message.id, ok: false, error: 'Preview layout failed' } : { id: message.id, ok: true,
        frame: commandFrame(PREVIEW_EPOCHS[previewPages.shift() ?? 'built']).payload.frame, documentPreview: true,
        caret: { frameEpoch: 1, caretRect: null }, selection: null, layoutRevision: 1 };
    } else if (message.type === 'fontRequirements') {
      response = { id: message.id, ok: true, requirementsJson: JSON.stringify(fontRequirements) };
    } else if (message.type === 'save') {
      response = { id: message.id, ok: true, saved: Uint8Array.of(4, 5).buffer, updates: [], version: 'saved' };
    } else if (message.type === 'documentRead') {
      response = { id: message.id, ok: true, read: { version: 'worker-version', value: message.read.kind === 'clipboardSlice' ? { text: 'Selected text', fragment: null }
        : message.read.kind === 'headings'
        ? [{ story: 'body', paraId: 'heading', text: 'Worker heading', level: 0, position: 1 }]
        : message.read.kind === 'exportStructuredWithPages' ? JSON.stringify({ ok: true, version: 'worker-version', content: { structured: { text: this.text }, layout: { pages: [{ pageIndex: 0 }, { pageIndex: 1 }] } } })
        : message.read.kind === 'sectionProperties' ? { version: 'worker-version', sections: [{ section: 'final', properties: { marginLeft: 2160 } }], paragraphs: [] }
        : message.read.kind === 'searchText' ? { matches: [], carried: -1 }
        : message.read.kind === 'rangeHighlights' ? null
        : message.read.kind === 'toolbarMetadata' ? { styles: [], fontFamilies: [] }
        : message.read.kind === 'navigationTarget' ? { loc: { story: 'body', paraId: 'heading', offset: 0 }, position: 1 }
        : message.read.kind === 'listStories' ? { ok: true, version: 'worker-version', stories: [{ story: 'hf:rId3', kind: 'header', root: 'hf:rId3' }] }
        : message.read.kind === 'readStories' ? { ok: true, version: 'worker-version', view: 'accepted', stories: [{ story: 'hf:rId3', ok: true, paragraphs: [] }], truncated: false }
        : { ok: true, view: 'accepted', version: 'worker-version', paragraphs: [] } } };
    } else if (message.type === 'commandInit') {
      const { session, incarnation, epoch, readOnly, layoutRevision } = message.request;
      const initial: ResidentCommandInitReply = {
        session, incarnation, epoch, version: 0, throughSeq: 0, outcomes: [], selection: null,
        editState: {
          readOnly, toolbar: null, effectiveFormatting: null, storedMarks: null,
          canUndo: false, canRedo: false, hyperlink: null, tableContext: null,
        },
        frame: {
          coversThroughSeq: 0, revision: 1,
          payload: { frame: commandFrame(1).payload.frame, caret: { frameEpoch: 1, caretRect: null }, layoutRevision,
            stagedPresentation: { frameEpoch: 1, pages: [] } },
        },
      };
      this.initial = initial;
      response = { id: message.id, ok: true, commandInitial: initial };
      if (holdInitial) { this.pendingReplies.push(response); return; }
    } else if (message.type === 'edit' && this.initial) {
      const command = message.command;
      const state = this.initial;
      const readOnlyRefusal = command.kind === 'insertText' && state.editState.readOnly;
      const refused = this.refuseNextEdit || readOnlyRefusal;
      this.refuseNextEdit = false;
      const cancelled = command.afterSeq !== undefined && this.refused.has(command.afterSeq);
      if (refused || cancelled) this.refused.add(command.seq);
      const readOnly = !refused && !cancelled && command.kind === 'setReadOnly'
        ? (command.payload as { readOnly: boolean }).readOnly : state.editState.readOnly;
      if (!refused && !cancelled && command.kind === 'insertText') this.text += (command.payload as { text: string }).text;
      const ack = {
        session: state.session, incarnation: state.incarnation, epoch: state.epoch,
        version: state.version + (!refused && !cancelled && command.kind === 'insertText' ? 1 : 0), throughSeq: command.seq,
        outcomes: [{
          seq: command.seq,
          status: refused ? 'refused' as const : cancelled ? 'cancelled' as const :
            command.kind === 'insertText' ? 'applied' as const : 'noop' as const,
          ...(refused ? { reason: readOnlyRefusal ? 'read-only' : 'stale-version' } :
            cancelled ? { reason: 'dependency-refused' } : {}),
        }],
        editState: { ...state.editState, readOnly },
        ...((command.kind === 'select' || (command.kind === 'navigate' && (command.payload as yrs.DocxCommandPayloads['navigate']).extend)) ? { selection: { anchor: { story: 'body', paraId: 'a', offset: 0 }, focus: { story: 'body', paraId: 'a', offset: 13 } } } : {}),
      };
      this.initial = { ...state, ...ack };
      response = { id: message.id, ok: true, commandAck: ack };
      if (!refused && !cancelled && command.kind === 'setFontEnvironment') {
        response.fontEnvironmentEvent = { session: state.session, incarnation: state.incarnation, epoch: state.epoch,
          phase: 'ready', revision: ++this.fontRevision, families: (command.payload as yrs.DocxCommandPayloads['setFontEnvironment']).families };
      }
      if (this.holdEdits) {
        this.pendingReplies.push(response);
        return;
      }
    } else {
      response = { id: message.id, ok: true };
      if (message.type === 'attachCanvases' && holdPreviewPaint) { this.pendingReplies.push(response); return; }
    }
    if (message.type === 'openSource' && message.source.cut && holdCut) { this.pendingReplies.push(response); return; }
    queueMicrotask(() => this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>));
  }
}

function fakeBinding(session: DocxWorkerEditSession<PeerlessEchoIntent>, initial: ResidentCommandInitReply): PeerlessEditSessionBinding {
  const snapshot: PeerlessViewSnapshot = {
    identity: { session: initial.session, incarnation: initial.incarnation, epoch: initial.epoch },
    version: initial.version, throughSeq: initial.throughSeq, presentedThroughSeq: 0,
    workerSelection: initial.selection, anchors: null, editState: initial.editState, geometry: null,
  };
  const viewStore: PeerlessEditSessionBinding['viewStore'] = {
    subscribe: () => () => {}, getSnapshot: () => snapshot, getServerSnapshot: () => snapshot,
    acceptWorkerState() {}, installFrameGeometry() {}, dispose() {},
  };
  const echoPainter: PeerlessEditSessionBinding['echoPainter'] = {
    installContext() {}, restoreBackground: () => false, preview: () => undefined,
    paint() {}, drop() {}, reconcile() {}, disable() {}, dispose() {},
  };
  const admission = createInputAdmission({ session, viewStore, echoPainter });
  return {
    session, viewStore, echoPainter, admission,
    commands: createPeerlessCommandBinding(admission),
    presentation: { hooks: { paint() {}, drop() {}, present() {} }, present: () => true, invalidate() {}, dispose() {} },
  };
}

function commandFrame(epoch: number, base = 0): DocxCommandFrame {
  const bytes = new Uint8Array(88);
  bytes.set([0x46, 0x44, 0x56, 0x31]);
  const data = new DataView(bytes.buffer);
  data.setUint16(4, 1, true);
  data.setUint16(6, 80, true);
  data.setUint32(8, bytes.length, true);
  data.setUint32(12, base === 0 ? 1 : 0, true);
  data.setBigUint64(16, 1n, true);
  data.setBigUint64(24, 1n, true);
  data.setBigUint64(32, BigInt(epoch), true);
  data.setBigUint64(40, BigInt(base), true);
  data.setUint32(56, 80, true);
  data.setUint32(60, 80, true);
  data.setUint32(64, 4, true);
  data.setUint32(68, 88, true);
  return {
    coversThroughSeq: epoch - 1, revision: epoch,
    ...(base ? { baseRevision: base } : {}),
    payload: { frame: bytes.buffer, caret: { frameEpoch: epoch, caretRect: null }, layoutRevision: epoch },
  };
}

function previewPage(epoch: number): frameRender.RetainedFramePage | null {
  if (epoch !== PREVIEW_EPOCHS.built) return null;
  return { pageIndex: 0, pageId: 1n, fingerprint: 1n, primitiveIds: new BigUint64Array(0),
    page: { pageIndex: 0, width: 816, height: 1056, primitives: [] } };
}

function installPreviewPages() {
  const apply = frameRender.applyFrameDelta;
  const applied = spyOn(frameRender, 'applyFrameDelta').mockImplementation((previous, delta) => {
    const frame = apply(previous, delta);
    const page = previous === null ? previewPage(delta.frameEpoch) : null;
    return page ? { ...frame, pages: [page] } : frame;
  });
  const canvas = window.HTMLCanvasElement.prototype as { transferControlToOffscreen?: () => unknown };
  const transfer = Object.getOwnPropertyDescriptor(canvas, 'transferControlToOffscreen');
  Object.defineProperty(canvas, 'transferControlToOffscreen', { configurable: true, value: () => ({}) });
  return () => {
    applied.mockRestore();
    if (transfer) Object.defineProperty(canvas, 'transferControlToOffscreen', transfer);
    else Reflect.deleteProperty(canvas, 'transferControlToOffscreen');
  };
}

function installCore() {
  return spyOn(core, 'useYrsCoreSession').mockReturnValue({
    session: null, sessionGeneration: null, replicaReady: false, workerProposalsReady: false,
    requestReplica() {}, storyBlocks: () => null, bodyBlocks: () => null,
    inputPositionMap: () => null, displayPositionToLoc: () => null, locToDisplayPosition: () => null,
    documentFromYrs: () => null, publishDirectInput() {}, previewing: false, handoffFrom: null,
    opening: false, notifyFramePresented() {}, failOpening: () => false,
    scheduleCompatibilityWarm() {}, cancelCompatibilityWarm() {},
  });
}

function installSurface() {
  return spyOn(paged, 'PeerlessPagedEditor').mockImplementation(function Surface({ sessionOptions, onReady }) {
    const [opened, setOpened] = useState(false);
    useEffect(() => {
      const controller = new AbortController();
      let session: DocxWorkerEditSession<PeerlessEchoIntent> | undefined;
      void sessionOptions.openConnection(sessionOptions.source, controller.signal).then((value) => {
        if (controller.signal.aborted) void value.client.dispose();
        else {
          session = peerless.createDocxWorkerEditSession<typeof sessionOptions.source, PeerlessEchoIntent>({
            source: sessionOptions.source,
            connection: value,
            echo: { paint() {}, drop() {}, present() {} },
            recovery: { reopen: () => Promise.reject(new Error('unexpected recovery')) },
          });
          binding = fakeBinding(session, value.initial);
          onReady?.(binding);
          setOpened(true);
        }
      }, (error: Error) => { if (!controller.signal.aborted) sessionOptions.onError?.(error); });
      return () => {
        void session?.dispose();
        controller.abort();
      };
    }, [sessionOptions]);
    return <div data-testid="peerless-paged-editor" data-opened={opened}>
      <div data-page-id="17"><canvas data-testid="page-canvas" /></div>
      <div data-page-index="0"><canvas data-testid="index-canvas" /></div>
    </div>;
  });
}

beforeEach(() => {
  holdInitial = false;
  holdPreviewPaint = false;
  failPreview = false;
  refuseCut = false;
  holdCut = false;
  failWholeOpen = false;
  previewPages = [];
  restorers.push(installPreviewPages());
  setMissingWorkerCapabilitiesForTests([]);
  requests.length = 0;
  workers.length = 0;
  binding = undefined;
  hostJson = JSON.stringify(hostMetadata());
  fontRequirements = [];
  let nextFontId = 70;
  const textEngine = spyOn(measurement, 'getRustTextEngine').mockResolvedValue({
    registerFont: () => nextFontId++, registerSubstituteFont: (base) => base, clearFonts() {},
  });
  restorers.push(() => textEngine.mockRestore());
  const worker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  Object.defineProperty(globalThis, 'Worker', { configurable: true, value: FakeWorker });
  restorers.push(() => {
    if (worker) Object.defineProperty(globalThis, 'Worker', worker);
    else Reflect.deleteProperty(globalThis, 'Worker');
  });
  coreSession = installCore();
  surface = installSurface();
  restorers.push(() => coreSession.mockRestore(), () => surface.mockRestore());
});

afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
  resetEngineChoiceForTests();
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a mounted host receives four proposals and a versioned plugin proposal event', async () => {
  const h = proposalHost();
  const ref = createRef<DocxEditorRef>();
  const events: { generation: string; version: string; previewVersion: number }[] = [];
  const snapshots: yrs.DocxProposalSnapshot[] = [];
  const plugin = defineDocxPlugin({ id: 'host-proposals', createState: () => null,
    async onEvent(context, event) {
      if (event.type !== 'proposal-change') return;
      events.push(event);
      snapshots.push(await ref.current!.getProposals());
      expect(context.snapshot.version).toBe(event.version);
    },
  });
  const readAt = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt').mockImplementation(h.readAt as never);
  restorers.push(() => readAt.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} ref={ref} allowHostProposals plugins={[plugin]} showToolbar={false} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { surface.mock.calls[0]![0].onReady!(h.binding); });
  await act(async () => {
    expect(await ref.current!.proposeChanges({ expectVersion: 'native-0', proposals: inputs })).toEqual({ ok: true, snapshot: h.snapshot() });
    expect(await ref.current!.getProposals()).toEqual(h.snapshot());
  });
  await waitFor(() => expect(snapshots.at(-1)?.proposals).toHaveLength(4));
  expect(events.at(-1)).toMatchObject({ generation: expect.any(String), version: h.snapshot().version, previewVersion: h.snapshot().previewVersion });
  view.unmount();
  h.controller.dispose();
});

test('a read-only viewer host proposing from its ref callback at mount gets the proposals once the worker document is ready', async () => {
  holdInitial = true;
  const h = proposalHost({ readOnly: true });
  const readAt = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt').mockImplementation(h.readAt as never);
  restorers.push(() => readAt.mockRestore());
  const errors: Error[] = [];
  let proposed: Promise<yrs.DocxProposalResult> | undefined;
  const attach = (editor: DocxEditorRef | null) => {
    if (editor && !proposed) proposed = editor.proposeChanges({ expectVersion: 'native-0', proposals: inputs });
  };
  const view = render(<DocxEditor {...peerlessProps} ref={attach} mode="viewing" readOnly allowHostProposals showToolbar={false}
    onError={error => errors.push(error)} />);
  await waitFor(() => expect(surface.mock.calls.length).toBeGreaterThan(0));
  expect(proposed).toBeDefined();
  expect(h.submit).not.toHaveBeenCalled();
  await act(async () => {
    surface.mock.calls.at(-1)![0].onReady!(h.binding);
    expect(await proposed).toEqual({ ok: true, snapshot: h.snapshot() });
  });
  expect(h.snapshot().proposals).toHaveLength(4);
  expect(errors).toEqual([]);
  view.unmount();
  h.controller.dispose();
});

test('worker host proposal sidebar cards follow the public visibility prop', async () => {
  const h = proposalHost();
  await h.submit('propose', { proposals: inputs }, {}).ack;
  h.binding.viewStore.getSnapshot().geometry!.pages = [0, 1].map(pageIndex => ({
    page: { pageId: String(pageIndex), pageIndex }, bounds: { x: 0, y: pageIndex * 820, width: 600, height: 800 },
    transform: [1, 0, 0, 1, 0, pageIndex * 820] as const,
  }));
  const readAt = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt').mockImplementation(h.readAt as never);
  restorers.push(() => readAt.mockRestore());
  const element = (showHostProposalsInSidebar: boolean) => <DocxEditor {...peerlessProps} allowHostProposals
    commentsSidebarOpen showHostProposalsInSidebar={showHostProposalsInSidebar} showToolbar={false} />;
  const view = render(element(false));
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { surface.mock.calls[0]![0].onReady!(h.binding); });
  expect(view.container.querySelectorAll('.docx-tracked-change-card')).toHaveLength(0);
  view.rerender(element(true));
  await waitFor(() => expect(view.container.querySelectorAll('.docx-tracked-change-card')).toHaveLength(4));
  expect(view.container.textContent).toContain('Reviewer');
  view.rerender(element(false));
  await waitFor(() => expect(view.container.querySelectorAll('.docx-tracked-change-card')).toHaveLength(0));
  view.unmount();
  h.controller.dispose();
});

test('worker host proposal cards alone leave an uncontrolled sidebar closed', async () => {
  const h = proposalHost();
  await h.submit('propose', { proposals: inputs }, {}).ack;
  h.binding.viewStore.getSnapshot().geometry!.pages = [0, 1].map(pageIndex => ({
    page: { pageId: String(pageIndex), pageIndex }, bounds: { x: 0, y: pageIndex * 820, width: 600, height: 800 },
    transform: [1, 0, 0, 1, 0, pageIndex * 820] as const,
  }));
  const readAt = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt').mockImplementation(h.readAt as never);
  restorers.push(() => readAt.mockRestore());
  const element = (open?: boolean) => <DocxEditor {...peerlessProps} allowHostProposals showHostProposalsInSidebar
    {...(open === undefined ? {} : { commentsSidebarOpen: open })} showToolbar={false} />;
  const view = render(element(true));
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { surface.mock.calls[0]![0].onReady!(h.binding); });
  await waitFor(() => expect(view.container.querySelectorAll('.docx-tracked-change-card')).toHaveLength(4));
  view.rerender(element());
  await act(async () => {});
  expect(view.container.querySelectorAll('.docx-tracked-change-card')).toHaveLength(0);
  view.unmount();
  h.controller.dispose();
});

test('worker shell draws unmanaged plugin overlays under the managed overlay layer', async () => {
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} pluginOverlays={<div data-testid="host-overlay" />} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const layer = view.getByTestId('host-overlay').parentElement!;
  const managed = view.getByTestId('plugin-overlays');
  expect(layer.parentElement).toBe(managed.parentElement);
  expect(Number(layer.style.zIndex)).toBeLessThan(Number(managed.style.zIndex));
});

test('unmanaged plugin sidebar items are placed from worker geometry without plugins', async () => {
  const h = proposalHost();
  h.binding.viewStore.getSnapshot().geometry!.pages = [{ page: { pageId: '0', pageIndex: 0 },
    bounds: { x: 0, y: 0, width: 600, height: 800 }, transform: [1, 0, 0, 1, 0, 0] as const }];
  const queries = { pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, width: 600, height: 800 }), anchorRect: () => null,
    rangeRects: (from: number) => from === 4 ? [{ pageIndex: 0, x: 10, y: 30, width: 5, height: 12 }] : [], dispose() {} } as unknown as frameRender.DisplayListQueries;
  const readAt = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt').mockImplementation(h.readAt as never);
  const load = spyOn(frameRender, 'loadRustDisplayListQueryEngine').mockResolvedValue({} as never);
  const create = spyOn(frameRender, 'createDisplayListQueries').mockReturnValue(queries);
  restorers.push(() => readAt.mockRestore(), () => load.mockRestore(), () => create.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false}
    pluginSidebarItems={[{ id: 'note', anchorPos: 4, render: () => <div data-testid="host-note" /> }]} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  view.getByTestId('peerless-paged-editor').setAttribute('data-peerless-pages-host', '');
  await act(async () => { surface.mock.calls[0]![0].onReady!(h.binding); });
  await waitFor(() => expect(view.getByTestId('host-note').parentElement!.style.top).toBe('30px'));
  view.unmount();
  h.controller.dispose();
});

test('plugin sidebar cards discard placement across source replacement and incarnations', async () => {
  const hosts = [proposalHost(), proposalHost(), proposalHost()];
  hosts[2]!.binding.viewStore.getSnapshot().identity.incarnation = 2;
  hosts[2]!.binding.viewStore.getSnapshot().geometry!.identity.incarnation = 2;
  for (const h of hosts) h.binding.viewStore.getSnapshot().geometry!.pages = [{ page: { pageId: '0', pageIndex: 0 },
    bounds: { x: 0, y: 0, width: 600, height: 800 }, transform: [1, 0, 0, 1, 0, 0] as const }];
  let y: number | null = 30;
  let currentHost = hosts[0]!;
  const readAt = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt')
    .mockImplementation(((...args: Parameters<typeof currentHost.readAt>) => currentHost.readAt(...args)) as never);
  const load = spyOn(frameRender, 'loadRustDisplayListQueryEngine').mockResolvedValue({} as never);
  const create = spyOn(frameRender, 'createDisplayListQueries').mockImplementation(() => {
    const top = y;
    return { pageBounds: (pageIndex: number) => ({ pageIndex, x: 0, y: 0, width: 600, height: 800 }),
      anchorRect: () => null, rangeRects: (from: number) => from === 4 && top !== null
        ? [{ pageIndex: 0, x: 10, y: top, width: 5, height: 12 }] : [], dispose() {} } as unknown as frameRender.DisplayListQueries;
  });
  restorers.push(() => readAt.mockRestore(), () => load.mockRestore(), () => create.mockRestore(),
    () => hosts.forEach(h => h.controller.dispose()));
  const items = [{ id: 'note', anchorPos: 4, render: () => <div data-testid="host-note" /> }];
  const element = (documentBuffer: Uint8Array) => <DocxEditor documentBuffer={documentBuffer} showToolbar={false} pluginSidebarItems={items} />;
  const view = render(element(source));
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  view.getByTestId('peerless-paged-editor').setAttribute('data-peerless-pages-host', '');
  await act(async () => { surface.mock.calls.at(-1)![0].onReady!(hosts[0]!.binding); });
  await waitFor(() => expect(view.getByTestId('host-note').parentElement!.style.top).toBe('30px'));
  const isPlaced = () => {
    const note = view.queryByTestId('host-note');
    const sidebar = note?.closest<HTMLElement>('.docx-unified-sidebar');
    return !!sidebar && sidebar.style.opacity !== '0' && note!.parentElement!.style.visibility !== 'hidden';
  };
  holdInitial = true;
  view.rerender(element(Uint8Array.of(4, 5, 6)));
  expect(isPlaced()).toBe(false);
  await waitFor(() => expect(workers.at(-1)!.pendingReplies.length).toBeGreaterThan(0));
  expect(isPlaced()).toBe(false);
  y = null;
  currentHost = hosts[1]!;
  await act(async () => {
    const worker = workers.at(-1)!;
    for (const response of worker.pendingReplies.splice(0)) worker.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  });
  await act(async () => { surface.mock.calls.at(-1)![0].onReady!(hosts[1]!.binding); });
  await waitFor(() => expect(create.mock.calls.length).toBeGreaterThan(1));
  expect(isPlaced()).toBe(false);
  y = 140;
  const frame = hosts[1]!.binding.viewStore.getSnapshot().geometry!;
  await act(async () => { hosts[1]!.binding.viewStore.installFrameGeometry({ ...frame, frameEpoch: frame.frameEpoch + 1 }); });
  await waitFor(() => expect(view.getByTestId('host-note').parentElement!.style.top).toBe('140px'));
  const previousQueries = create.mock.calls.length;
  y = null;
  currentHost = hosts[2]!;
  await act(async () => { surface.mock.calls.at(-1)![0].onReady!(hosts[2]!.binding); });
  await waitFor(() => expect(create.mock.calls.length).toBeGreaterThan(previousQueries));
  expect(isPlaced()).toBe(false);
  y = 250;
  const recoveredFrame = hosts[2]!.binding.viewStore.getSnapshot().geometry!;
  await act(async () => { hosts[2]!.binding.viewStore.installFrameGeometry({ ...recoveredFrame, frameEpoch: recoveredFrame.frameEpoch + 1 }); });
  await waitFor(() => expect(view.getByTestId('host-note').parentElement!.style.top).toBe('250px'));
});

test('unmanaged plugin overlays wait for a ready document and clear on replacement', async () => {
  holdInitial = true;
  const element = (documentBuffer: Uint8Array) => <DocxEditor documentBuffer={documentBuffer} showToolbar={false}
    pluginOverlays={<div data-testid="host-overlay" />} />;
  const view = render(element(source));
  expect(view.queryByTestId('host-overlay')).toBeNull();
  await waitFor(() => expect(workers[0]!.pendingReplies.length).toBeGreaterThan(0));
  expect(view.queryByTestId('host-overlay')).toBeNull();
  await act(async () => {
    for (const response of workers[0]!.pendingReplies.splice(0)) workers[0]!.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  });
  await waitFor(() => expect(view.queryByTestId('host-overlay')).not.toBeNull());
  view.rerender(element(Uint8Array.of(4, 5, 6)));
  expect(view.queryByTestId('host-overlay')).toBeNull();
});

test('unmanaged plugin overlays stay hidden for placeholder and parse errors', async () => {
  const element = (documentBuffer?: Uint8Array) => <DocxEditor documentBuffer={documentBuffer} showToolbar={false}
    pluginOverlays={<div data-testid="host-overlay" />} />;
  const view = render(element());
  expect(view.queryByTestId('host-overlay')).toBeNull();
  failWholeOpen = true;
  const onError = mock(() => {});
  view.rerender(<DocxEditor {...peerlessProps} showToolbar={false} onError={onError}
    pluginOverlays={<div data-testid="host-overlay" />} />);
  await waitFor(() => expect(onError).toHaveBeenCalled());
  expect(view.queryByTestId('host-overlay')).toBeNull();
});

test('unmanaged plugin sidebar items keyed to a host proposal card are placed at that card', async () => {
  const h = proposalHost();
  await h.submit('propose', { proposals: inputs }, {}).ack;
  h.binding.viewStore.getSnapshot().geometry!.pages = [0, 1].map(pageIndex => ({
    page: { pageId: String(pageIndex), pageIndex }, bounds: { x: 0, y: pageIndex * 820, width: 600, height: 800 },
    transform: [1, 0, 0, 1, 0, pageIndex * 820] as const,
  }));
  const readAt = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt').mockImplementation(h.readAt as never);
  restorers.push(() => readAt.mockRestore());
  const anchorKey = `revision-${frameRender.yrsIdToNumericId('revision-change-1')}`;
  const view = render(<DocxEditor {...peerlessProps} allowHostProposals commentsSidebarOpen showHostProposalsInSidebar showToolbar={false}
    pluginSidebarItems={[{ id: 'note', anchorPos: -1, anchorKey, render: () => <div data-testid="host-note" /> }]} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { surface.mock.calls[0]![0].onReady!(h.binding); });
  await waitFor(() => expect(view.container.querySelectorAll('.docx-tracked-change-card')).toHaveLength(4));
  const note = view.getByTestId('host-note').parentElement!;
  await waitFor(() => expect([note.style.position, note.style.visibility]).toEqual(['absolute', '']));
  const cards = [...view.container.querySelectorAll('.docx-tracked-change-card')]
    .map(card => Number.parseFloat(card.closest<HTMLElement>('.docx-unified-sidebar > div > div')!.style.top)).sort((a, b) => a - b);
  expect(Number.parseFloat(note.style.top)).toBeGreaterThan(cards[0]!);
  expect(Number.parseFloat(note.style.top)).toBeLessThan(cards[1]!);
  view.unmount();
  h.controller.dispose();
});

test('d8a_peerless_mount_creates_no_main_document', async () => {
  const main = spyOn(yrs, 'createYrsSession').mockImplementation(() => {
    throw new Error('main document created');
  });
  const renderer = spyOn(display, 'useCanvasRenderer');
  const workerSession = spyOn(peerless, 'createResidentCommandConnection');
  const commands = spyOn(peerless, 'createDocxWorkerEditSession');
  restorers.push(() => main.mockRestore(), () => renderer.mockRestore(), () => workerSession.mockRestore(),
    () => commands.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} initialZoom={1.25} />);
  expect(core.useYrsCoreSession).not.toHaveBeenCalled();
  expect(main).not.toHaveBeenCalled();
  expect(renderer).not.toHaveBeenCalled();
  await waitFor(() => expect(paged.PeerlessPagedEditor).toHaveBeenCalled());
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  expect(workerSession).toHaveBeenCalledTimes(1);
  expect(commands).toHaveBeenCalledTimes(1);
  expect(workers).toHaveLength(1);
  expect(requests.filter(({ type }) => type === 'openSource')).toHaveLength(1);
  expect(requests.filter(({ type }) => type === 'commandInit')).toHaveLength(1);
  expect(requests.some(({ type }) => type === 'encodeState' || type === 'bootstrap')).toBe(false);
  const props = surface.mock.calls[0][0];
  expect(props.zoom).toBe(1.25);
  expect(props.sessionOptions.source).toEqual(source);
  view.unmount();
  await waitFor(() => expect(workers[0].terminate).toHaveBeenCalledTimes(1));
});

test('parsed input opens once in the worker with source options and no main engine', async () => {
  const parsed = yrs.decodeDocxHostJson(hostJson, source).document;
  const view = render(<DocxEditor document={parsed} mediaTokens previewFirstPage showToolbar={false} />);
  await waitFor(() => expect(requests.some(request => request.type === 'openSource')).toBe(true));
  expect(requests.find(request => request.type === 'openSource')).toMatchObject({ source: {
    document: parsed, mediaTokens: true,
  } });
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  expect(requests.filter(request => request.type === 'commandInit')).toHaveLength(1);
  expect(requests.filter(request => request.type === 'previewSource')).toHaveLength(1);
  expect(requests.some(request => request.type === 'open' || request.type === 'bootstrap')).toBe(false);
  expect(requests.filter(request => request.type === 'warm')).toHaveLength(1);
  expect(coreSession).not.toHaveBeenCalled();
});

test('previewFirstPage paints the first page with the command fonts and layout before the command session lays out', async () => {
  render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} />);
  await waitFor(() => expect(requests.some(request => request.type === 'commandInit')).toBe(true));
  const opened = requests.find(request => request.type === 'openSource');
  const preview = requests.find(request => request.type === 'previewSource');
  const init = requests.find(request => request.type === 'commandInit');
  if (opened?.type !== 'openSource' || preview?.type !== 'previewSource' || init?.type !== 'commandInit') throw new Error('Worker did not preview');
  expect(opened.source.preview).toBeUndefined();
  const order = requests.map(request => request.type);
  expect(order.indexOf('fontRequirements')).toBeLessThan(order.indexOf('previewSource'));
  expect(order.indexOf('previewSource')).toBeLessThan(order.indexOf('attachCanvases'));
  expect(order.indexOf('attachCanvases')).toBeLessThan(order.indexOf('commandInit'));
  const { layoutInput, ...options } = preview.preview;
  expect(options).toEqual({ fonts: init.request.fonts, fontRegistrationSlots: init.request.fontRegistrationSlots, extras: init.request.extras });
  expect(JSON.parse(layoutInput)).toEqual({ ...JSON.parse(init.request.layoutInput), cachedPageTotals: true });
  expect(JSON.parse(init.request.layoutInput)).not.toHaveProperty('cachedPageTotals');
});

test('previewFirstPage previews a cut of the source before the whole source opens', async () => {
  render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} />);
  await waitFor(() => expect(requests.some(request => request.type === 'commandInit')).toBe(true));
  const order = requests.map(request => request.type).filter(type => ['openSource', 'previewSource', 'attachCanvases', 'commandInit'].includes(type));
  expect(order.slice(0, order.indexOf('commandInit') + 1)).toEqual(['openSource', 'previewSource', 'attachCanvases', 'openSource', 'commandInit']);
  const [cut, whole] = requests.filter(request => request.type === 'openSource');
  if (cut?.type !== 'openSource' || whole?.type !== 'openSource') throw new Error('Worker did not open the source');
  expect(cut.source.cut).toEqual({ blocks: 200, paragraphBudget: 256 });
  expect(whole.source.cut).toBeUndefined();
  expect(Buffer.from(cut.source.bytes!)).toEqual(Buffer.from(whole.source.bytes!));
});

test('a declined source cut previews the whole source before the command session', async () => {
  refuseCut = true;
  render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} />);
  await waitFor(() => expect(requests.some(request => request.type === 'commandInit')).toBe(true));
  const order = requests.map(request => request.type).filter(type => ['openSource', 'previewSource', 'attachCanvases', 'commandInit'].includes(type));
  expect(order.slice(0, order.indexOf('commandInit') + 1)).toEqual(['openSource', 'openSource', 'previewSource', 'attachCanvases', 'commandInit']);
});

test('the preview releases its worker surfaces once the command session presents', async () => {
  const view = render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await waitFor(() => expect(requests.filter(request => request.type === 'attachCanvases')).toHaveLength(2));
  const order = requests.map(request => request.type);
  const [painted, released] = requests.filter(request => request.type === 'attachCanvases');
  if (painted?.type !== 'attachCanvases') throw new Error('The preview attached no surfaces');
  expect(painted.pages.length).toBeGreaterThan(0);
  expect(order.indexOf('attachCanvases')).toBeLessThan(order.indexOf('commandInit'));
  expect(order.lastIndexOf('attachCanvases')).toBeGreaterThan(order.indexOf('commandInit'));
  expect(released).toMatchObject({ pages: [], activePageIds: [] });
  expect(view.queryByLabelText('Document preview')).toBeNull();
});

test('a cut preview without a built first page keeps opening and previews the whole source', async () => {
  holdInitial = true;
  previewPages = ['none', 'built'];
  const view = render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} loadingIndicator={<span>Opening source</span>} />);
  await waitFor(() => expect(requests.some(request => request.type === 'commandInit')).toBe(true));
  const order = requests.map(request => request.type).filter(type => ['openSource', 'previewSource', 'attachCanvases', 'commandInit'].includes(type));
  expect(order).toEqual(['openSource', 'previewSource', 'openSource', 'previewSource', 'attachCanvases', 'commandInit']);
  expect(view.queryByText('Opening source')).toBeNull();
  expect(view.getByLabelText('Document preview')).toBeTruthy();
});

test('the preview budget covers the source cut and its late steps preview nothing', async () => {
  holdCut = true;
  const budget = peerlessPreviewPaint.timeoutMs;
  peerlessPreviewPaint.timeoutMs = 4_321;
  restorers.push(() => { peerlessPreviewPaint.timeoutMs = budget; });
  const schedule = globalThis.setTimeout;
  const expiries: (() => void)[] = [];
  const timers = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: () => void, delay?: number, ...rest: unknown[]) => {
    if (delay !== 4_321) return schedule(callback, delay, ...rest);
    expiries.push(callback);
    return schedule(() => {}, 0);
  }) as typeof setTimeout);
  restorers.push(() => timers.mockRestore());
  const turns = async (done: () => boolean) => {
    for (let turn = 0; turn < 200 && !done(); turn += 1) await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
    expect(done()).toBe(true);
  };
  const view = render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} />);
  await turns(() => expiries.length === 1);
  expect(requests.filter(request => request.type === 'openSource')).toHaveLength(1);
  act(() => expiries[0]!());
  await turns(() => view.queryByTestId('peerless-paged-editor')?.dataset.opened === 'true');
  await act(async () => workers[0].onmessage?.({ data: workers[0].pendingReplies.shift()! } as MessageEvent<ResidentEngineWorkerResponse>));
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  expect(requests.map(request => request.type).filter(type => ['openSource', 'previewSource', 'attachCanvases', 'commandInit'].includes(type)))
    .toEqual(['openSource', 'openSource', 'commandInit']);
});

test('a whole source that fails after the cut painted keeps the preview and reports the failure', async () => {
  failWholeOpen = true;
  const errors: Error[] = [];
  const view = render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} onError={error => errors.push(error)}
    loadingIndicator={<span>Opening source</span>} />);
  await waitFor(() => expect(errors.map(error => error.message)).toEqual(['Whole source failed to open']));
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  expect(view.getByLabelText('Document preview')).toBeTruthy();
  expect(view.queryByText('Opening source')).toBeNull();
  expect(workers[0].terminate).not.toHaveBeenCalled();
  expect(requests.some(request => request.type === 'commandInit')).toBe(false);
});

test('a buffer load whose whole source fails after the cut painted rejects, and a later load opens', async () => {
  const ref = createRef<DocxEditorRef>();
  const view = render(<DocxEditor {...peerlessProps} ref={ref} previewFirstPage showToolbar={false} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  failWholeOpen = true;
  let outcome = 'pending';
  act(() => { ref.current!.loadDocumentBuffer(Uint8Array.of(8, 9)).catch((error: Error) => { outcome = error.message; }); });
  await waitFor(() => expect(outcome).toBe('Whole source failed to open'));
  await expect(ref.current!.whenLayoutComplete()).rejects.toThrow('Whole source failed to open');
  failWholeOpen = false;
  let pending!: Promise<void>;
  act(() => { pending = ref.current!.loadDocumentBuffer(Uint8Array.of(10, 11)); });
  await pending;
});

async function paintPreviewAfterBudget(onFirstPagePainted?: () => void) {
  holdInitial = true;
  holdPreviewPaint = true;
  const budget = peerlessPreviewPaint.timeoutMs;
  peerlessPreviewPaint.timeoutMs = 4_321;
  const schedule = globalThis.setTimeout;
  const expiries: (() => void)[] = [];
  globalThis.setTimeout = ((callback: () => void, delay?: number, ...rest: unknown[]) =>
    delay === 4_321 ? (expiries.push(callback), schedule(() => {}, 0)) : schedule(callback, delay, ...rest)) as typeof setTimeout;
  restorers.push(() => { peerlessPreviewPaint.timeoutMs = budget; globalThis.setTimeout = schedule; });
  const view = render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} onFirstPagePainted={onFirstPagePainted}
    loadingIndicator={<span>Opening source</span>} />);
  await waitFor(() => expect(requests.some(request => request.type === 'attachCanvases')).toBe(true));
  act(() => expiries[0]!());
  await waitFor(() => expect(requests.some(request => request.type === 'commandInit')).toBe(true));
  await act(async () => workers[0].onmessage?.({ data: workers[0].pendingReplies.shift()! } as MessageEvent<ResidentEngineWorkerResponse>));
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  return view;
}

test('a preview that paints after its budget expired presents nothing', async () => {
  const painted = mock(() => {});
  const view = await paintPreviewAfterBudget(painted);
  expect(painted).not.toHaveBeenCalled();
  expect(view.getByText('Opening source')).toBeTruthy();
});

test('a preview that paints after its budget expired settles and releases its surfaces before the document opens', async () => {
  const settled = spyOn(AbortSignal.prototype, 'removeEventListener');
  restorers.push(() => settled.mockRestore());
  const view = await paintPreviewAfterBudget();
  expect(settled).toHaveBeenCalledWith('abort', expect.any(Function));
  for (const reply of workers[0].pendingReplies.splice(0)) await act(async () => workers[0].onmessage?.({ data: reply } as MessageEvent<ResidentEngineWorkerResponse>));
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  expect(requests.filter(request => request.type === 'attachCanvases'))
    .toMatchObject([{ pages: [expect.anything()] }, { pages: [], activePageIds: [] }]);
});

for (const failure of ['layout', 'paint'] as const) {
  test(`a preview ${failure} failure still opens the document`, async () => {
    const errors: Error[] = [];
    failPreview = failure === 'layout';
    holdPreviewPaint = failure === 'paint';
    const budget = peerlessPreviewPaint.timeoutMs;
    peerlessPreviewPaint.timeoutMs = 20;
    restorers.push(() => { peerlessPreviewPaint.timeoutMs = budget; });
    const view = render(<DocxEditor {...peerlessProps} previewFirstPage showToolbar={false} onError={error => errors.push(error)} />);
    await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
    expect(requests.filter(request => request.type === 'commandInit')).toHaveLength(1);
    expect(errors.map(error => error.message)).toEqual(failure === 'layout' ? ['Preview layout failed'] : []);
  });
}

test('preview paint hides opening while commands and plugins wait for handoff', async () => {
  holdInitial = true;
  holdPreviewPaint = true;
  const ref = createRef<DocxEditorRef>();
  const initialize = mock(() => {});
  const plugin = defineDocxPlugin({ id: 'preview-check', createState: () => null, initialize });
  const view = render(<DocxEditor {...peerlessProps} ref={ref} previewFirstPage plugins={[plugin]}
    showToolbar={false} loadingIndicator={<span>Opening source</span>} />);
  await waitFor(() => expect(requests.some(request => request.type === 'attachCanvases')).toBe(true));
  expect(view.getByText('Opening source')).toBeTruthy();
  expect(ref.current!.commands.getState('save').enabled).toBe(false);
  expect(initialize).not.toHaveBeenCalled();
  await act(async () => workers[0].onmessage?.({ data: workers[0].pendingReplies.shift()! } as MessageEvent<ResidentEngineWorkerResponse>));
  await waitFor(() => expect(requests.some(request => request.type === 'commandInit')).toBe(true));
  expect(view.queryByText('Opening source')).toBeNull();
  expect(view.getByLabelText('Document preview')).toBeTruthy();
  expect(initialize).not.toHaveBeenCalled();
  await act(async () => workers[0].onmessage?.({ data: workers[0].pendingReplies.shift()! } as MessageEvent<ResidentEngineWorkerResponse>));
  await waitFor(() => expect(initialize).toHaveBeenCalledTimes(1));
  expect(view.queryByLabelText('Document preview')).toBeNull();
  expect(ref.current!.commands.getState('save').enabled).toBe(true);
  expect(requests.filter(request => request.type === 'openSource').map(request => request.type === 'openSource' && !!request.source.cut)).toEqual([true, false]);
});

test('worker render props initialize layout and update through the shared command tail', async () => {
  const theme = { colorScheme: { accent1: 'FF0000' } } as DocxEditorProps['theme'];
  const ref = createRef<DocxEditorRef>();
  const view = render(<DocxEditor {...peerlessProps} ref={ref} theme={theme} colorMode="dark" showHiddenText showTableGridlines readOnly />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const opened = requests.find((request) => request.type === 'commandInit');
  if (opened?.type !== 'commandInit') throw new Error('Worker did not initialize');
  expect(JSON.parse(opened.request.layoutInput).renderEnv).toMatchObject({ themeColors: { accent1: 'FF0000' }, showHiddenText: true, showTableGridlines: true });
  expect(ref.current!.commands.getState('tableGridlines').active).toBe(true);
  expect(view.getByTestId('docx-editor').classList.contains('dark')).toBe(true);
  view.rerender(<DocxEditor {...peerlessProps} ref={ref} theme={{ colorScheme: { accent1: '0000FF' } } as DocxEditorProps['theme']}
    colorMode="light" showHiddenText={false} showTableGridlines={false} readOnly />);
  await waitFor(() => expect(requests.filter((request) => request.type === 'edit' && request.command.kind === 'setRenderEnvironment')).toHaveLength(1));
  expect(requests.find((request) => request.type === 'edit' && request.command.kind === 'setRenderEnvironment')).toMatchObject({ command: {
    payload: { themeColors: { accent1: '0000FF' }, showHiddenText: false, showTableGridlines: false },
  } });
  expect(view.getByTestId('docx-editor').classList.contains('dark')).toBe(false);
  expect(requests.filter((request) => request.type === 'openSource')).toHaveLength(1);
  expect(coreSession).not.toHaveBeenCalled();
});

test('font callbacks follow worker readiness and a reverted provider is reapplied after a held acknowledgement', async () => {
  const onFontsLoaded = mock(() => {});
  const view = render(<DocxEditor {...peerlessProps} onFontsLoaded={onFontsLoaded} />);
  await waitFor(() => expect(onFontsLoaded).toHaveBeenCalledTimes(1));
  const client = workers[0];
  client.holdEdits = true;
  view.rerender(<DocxEditor {...peerlessProps} onFontsLoaded={onFontsLoaded} measurementFontProvider={{ resolve: () => undefined }} />);
  await waitFor(() => expect(client.pendingReplies).toHaveLength(1));
  expect(onFontsLoaded).toHaveBeenCalledTimes(1);
  view.rerender(<DocxEditor {...peerlessProps} onFontsLoaded={onFontsLoaded} />);
  client.holdEdits = false;
  await act(async () => {
    for (const response of client.pendingReplies.splice(0)) client.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  });
  await waitFor(() => expect(requests.filter(request => request.type === 'edit' && request.command.kind === 'setFontEnvironment')).toHaveLength(3));
  await waitFor(() => expect(onFontsLoaded).toHaveBeenCalledTimes(3));
  expect(requests.filter(request => request.type === 'openSource')).toHaveLength(1);
  expect(coreSession).not.toHaveBeenCalled();
});

test('custom fonts fetch bytes for worker faces without global CSS and wait for composition', async () => {
  const original = document.head.innerHTML;
  const fontBytes = new Uint8Array(16);
  new DataView(fontBytes.buffer).setUint32(0, 0x774f4632);
  const fetchFont = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(fontBytes));
  const measure = spyOn(measurement, 'createRustMeasureSource').mockImplementation(({ engine }) => {
    let embedded: { data: ArrayBuffer }[] = [];
    let slots: number[] = [];
    return { setCompat() {}, setEmbeddedFaces(faces) { embedded = faces; },
      async prepareFontRequirements() { slots = embedded.map(face => engine.registerFont(new Uint8Array(face.data))); return true; },
      measurementConfigForRequirements: () => ({ fontChains: { 'custom|0|0': slots }, defaults: { fontFamily: 'Custom', fontSize: 11 },
        compat: { noLeading: false, doNotExpandShiftReturn: false }, authoritativeShaping: true }), clear() {} };
  });
  restorers.push(() => fetchFont.mockRestore(), () => measure.mockRestore());
  const onFontsLoaded = mock(() => {});
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const reservation = binding!.admission.reserveComposition();
  view.rerender(<DocxEditor {...peerlessProps} fonts={[{ family: 'Custom', src: '/custom.woff2' }]} onFontsLoaded={onFontsLoaded} />);
  await act(async () => { await Promise.resolve(); });
  expect(fetchFont).not.toHaveBeenCalled();
  await act(async () => binding!.admission.settleComposition(reservation.slot, { text: 'x', target: reservation.target }));
  await waitFor(() => expect(onFontsLoaded).toHaveBeenCalledTimes(1));
  const fontCommand = requests.find(request => request.type === 'edit' && request.command.kind === 'setFontEnvironment');
  expect(fontCommand).toMatchObject({ command: { payload: { faces: [{ family: 'Custom', data: fontBytes }],
    chains: { 'custom|0|0': [0] }, families: ['Custom'] } } });
  expect(document.head.innerHTML).toBe(original);
});

test('peerless shell ref and plugin readiness use worker state', async () => {
  const ref = createRef<DocxEditorRef>();
  const setup = mock(() => {});
  const plugin = defineDocxPlugin({ id: 'shell-check', createState: () => null, initialize: setup });
  const view = render(<DocxEditor {...peerlessProps} ref={ref} plugins={[plugin]} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  expect(ref.current).not.toBeNull();
  expect(() => ref.current!.getDocument()).toThrow(LEGACY_SYNC_API_PROP);
  expect(() => ref.current!.onContentChange(() => {})).toThrow(LEGACY_SYNC_API_PROP);
  await waitFor(() => expect(setup).toHaveBeenCalledTimes(1));
  expect(coreSession).not.toHaveBeenCalled();
});

test('default rejects Document callbacks through an alert and one error without throwing', () => {
  const onError = mock((_error: Error) => {});
  const invalid = { onChange() {} } as unknown as DocxEditorProps;
  const view = render(<DocxEditor {...peerlessProps} {...invalid} onError={onError} />);
  expect(view.getByRole('alert').textContent).toContain(LEGACY_SYNC_API_PROP);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls[0]).toEqual([expect.any(PeerlessSyncApiError)]);
  expect(String(onError.mock.calls[0])).toContain('onChange(Document)');
  view.rerender(<DocxEditor {...peerlessProps} {...invalid} onError={onError} documentName="Changed" />);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(coreSession).not.toHaveBeenCalled();
  expect(workers).toHaveLength(0);
});

test('peerless plugin story reads go to the worker', async () => {
  let context: DocxPluginContext<null> | undefined;
  const plugin = defineDocxPlugin({ id: 'stories', createState: () => null, initialize: (value) => { context = value; } });
  const view = render(<DocxEditor {...peerlessProps} plugins={[plugin]} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await waitFor(() => expect(context).toBeDefined());
  expect(await context!.read.listStories()).toEqual({
    ok: true, version: 'worker-version', stories: [{ story: 'hf:rId3', kind: 'header', root: 'hf:rId3' }],
  });
  const request = { stories: ['header'], view: 'accepted' } as const;
  expect(await context!.read.readStories(request)).toMatchObject({ ok: true, stories: [{ story: 'hf:rId3', ok: true }] });
  const reads = requests.flatMap((entry) => entry.type === 'documentRead' ? [entry.read] : []);
  expect(reads).toContainEqual({ kind: 'listStories' });
  expect(reads).toContainEqual({ kind: 'readStories', request });
  expect(coreSession).not.toHaveBeenCalled();
});

test('peerless rejects Document callbacks with the sync opt-in named', () => {
  const onError = mock((_error: Error) => {});
  const invalid = { onChange() {} } as unknown as DocxEditorProps;
  const view = render(<DocxEditor {...peerlessProps} {...invalid} onError={onError} />);
  expect(view.getByRole('alert').textContent).toContain(LEGACY_SYNC_API_PROP);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onError.mock.calls[0]).toEqual([expect.any(PeerlessSyncApiError)]);
  expect(String(onError.mock.calls[0])).toContain('onChange(Document)');
  expect(coreSession).not.toHaveBeenCalled();
  expect(workers).toHaveLength(0);
});

test('peerless zoom controls stay usable while opening and update the worker surface', async () => {
  const ref = createRef<DocxEditorRef>();
  const view = render(<DocxEditor {...peerlessProps} ref={ref} initialZoom={1.25} />);
  await waitFor(() => expect(ref.current).not.toBeNull());
  expect(ref.current!.commands.getState('zoom').enabled).toBe(true);
  await act(async () => { expect(await ref.current!.commands.execute('zoom', { scale: 1.5 })).toMatchObject({ ok: true }); });
  await waitFor(() => expect(surface.mock.calls.at(-1)![0].zoom).toBe(1.5));
  expect(ref.current!.getZoom()).toBe(1.5);
  expect(workers).toHaveLength(1);
  expect(view.getByTestId('docx-editor')).toBeTruthy();
});

test('peerless shell keeps the opening indicator until worker readiness and supplies an empty placeholder', async () => {
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} loadingIndicator={<span>Opening source</span>} />);
  expect(view.getByText('Opening source')).toBeTruthy();
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  expect(view.queryByText('Opening source')).toBeNull();
  view.unmount();
  const empty = render(<DocxEditor showToolbar={false} />);
  await waitFor(() => expect(empty.getByText('No document loaded')).toBeTruthy());
  empty.rerender(<DocxEditor showToolbar={false} placeholder={<span>Choose source</span>} />);
  expect(empty.getByText('Choose source')).toBeTruthy();
});

test('shell background returns focus to the worker input', async () => {
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const host = view.getByTestId('peerless-paged-editor').parentElement!;
  const input = document.createElement('textarea');
  input.className = 'paged-editor__peerless-input';
  host.appendChild(input);
  const focus = spyOn(input, 'focus');
  restorers.push(() => focus.mockRestore());
  act(() => { host.parentElement!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); });
  expect(focus).toHaveBeenCalled();
});

test('worker shell save flushes pending edits and calls host callbacks without a main document', async () => {
  const ref = createRef<DocxEditorRef>();
  const onSave = mock(() => {});
  const onSaveRequest = mock(async () => true);
  const view = render(<DocxEditor {...peerlessProps} ref={ref} onSave={onSave} onSaveRequest={onSaveRequest} downloadOnSave={false} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const result = await ref.current!.save();
  expect(result).toEqual(Uint8Array.of(4, 5).buffer);
  expect(onSave).toHaveBeenCalledTimes(1);
  await act(async () => { expect(await ref.current!.commands.execute('save', null)).toMatchObject({ ok: true }); });
  expect(onSaveRequest).toHaveBeenCalledTimes(1);
  expect(onSave).toHaveBeenCalledTimes(2);
  expect(requests.filter(({ type }) => type === 'save')).toHaveLength(2);
  expect(coreSession).not.toHaveBeenCalled();
});

test('a failed worker save command reports a recoverable error and allows retry', async () => {
  const ref = createRef<DocxEditorRef>();
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  const failure = new Error('Save request failed');
  let fail = true;
  const view = render(<DocxEditor {...peerlessProps} ref={ref} onError={onError} downloadOnSave={false}
    onSaveRequest={() => { if (fail) throw failure; return true; }} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { expect(await ref.current!.commands.execute('save', null)).toMatchObject({ ok: false }); });
  expect(onError.mock.calls).toEqual([[failure, { recoverable: true }]]);
  expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true');
  fail = false;
  await act(async () => { expect(await ref.current!.commands.execute('save', null)).toMatchObject({ ok: true }); });
});

test('worker saves retain source comments unless the host controls them', async () => {
  const metadata = hostMetadata();
  const comments = [{ id: 1, author: 'Writer', content: [] }];
  Object.assign(metadata.envelope.document.package.document, { comments });
  hostJson = JSON.stringify(metadata);
  const ref = createRef<DocxEditorRef>();
  const view = render(<DocxEditor {...peerlessProps} ref={ref} downloadOnSave={false} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { await ref.current!.save(); });
  expect(requests.filter((request) => request.type === 'save').at(-1)).toMatchObject({ comments });
  view.rerender(<DocxEditor {...peerlessProps} ref={ref} downloadOnSave={false} comments={[]} />);
  await act(async () => { await ref.current!.save(); });
  expect(requests.filter((request) => request.type === 'save').at(-1)).toMatchObject({ comments: [] });
  expect(coreSession).not.toHaveBeenCalled();
});

test('worker ref buffer loads replace the worker and resolve on readiness', async () => {
  const ref = createRef<DocxEditorRef>();
  const view = render(<DocxEditor {...peerlessProps} ref={ref} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  let pending!: Promise<void>;
  act(() => { pending = ref.current!.loadDocumentBuffer(Uint8Array.of(8, 9)); });
  await waitFor(() => expect(workers).toHaveLength(2));
  await waitFor(() => expect(ref.current!.commands.getState('save').enabled).toBe(true));
  await pending;
  expect(requests.filter(({ type }) => type === 'openSource')).toHaveLength(2);
  await waitFor(() => expect(workers[0]!.terminate).toHaveBeenCalled());
  expect(coreSession).not.toHaveBeenCalled();
});

test('worker ref document loads open the parsed document in a new worker', async () => {
  const ref = createRef<DocxEditorRef>();
  const parsed = yrs.decodeDocxHostJson(hostJson, source).document;
  const view = render(<DocxEditor {...peerlessProps} ref={ref} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  act(() => { expect(ref.current!.loadDocument(parsed)).toBeUndefined(); });
  await waitFor(() => expect(workers).toHaveLength(2));
  await waitFor(() => expect(requests.filter(({ type }) => type === 'openSource')).toHaveLength(2));
  expect(requests.filter(request => request.type === 'openSource').at(-1)).toMatchObject({ source: { document: parsed } });
  await waitFor(() => expect(ref.current!.commands.getState('save').enabled).toBe(true));
  await waitFor(() => expect(workers[0]!.terminate).toHaveBeenCalled());
  expect(coreSession).not.toHaveBeenCalled();
});

test('worker file picker delegates to the host and save can be intercepted', async () => {
  const ref = createRef<DocxEditorRef>();
  const onOpen = mock(async (_file: File) => {});
  const view = render(<DocxEditor {...peerlessProps} ref={ref} onOpen={onOpen} onSaveRequest={() => {}} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const file = new File([source], 'picked.docx');
  fireEvent.change(view.getByTestId('docx-file-input'), { target: { files: [file] } });
  await waitFor(() => expect(onOpen).toHaveBeenCalledWith(file));
  await act(async () => { expect(await ref.current!.commands.execute('save', null)).toMatchObject({ ok: true }); });
  expect(requests.filter(({ type }) => type === 'save')).toHaveLength(0);
  expect(workers).toHaveLength(1);
});

test('a throwing file-open callback reports a recoverable error and keeps the document open', async () => {
  const failure = new Error('Open request failed');
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  const view = render(<DocxEditor {...peerlessProps} onError={onError} onOpen={() => { throw failure; }} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const editor = view.getByTestId('peerless-paged-editor');
  await act(async () => {
    fireEvent.change(view.getByTestId('docx-file-input'), { target: { files: [new File([source], 'picked.docx')] } });
  });
  expect(onError.mock.calls).toEqual([[failure, { recoverable: true }]]);
  expect(view.getByTestId('peerless-paged-editor')).toBe(editor);
  expect(editor.dataset.opened).toBe('true');
});

test('a superseded file open reports a recoverable error while the replacement opens', async () => {
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  const read = spyOn(utils, 'toArrayBuffer').mockResolvedValue(source.buffer);
  restorers.push(() => read.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} onError={onError} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => {
    const input = view.getByTestId('docx-file-input');
    fireEvent.change(input, { target: { files: [new File([source], 'first.docx')] } });
    fireEvent.change(input, { target: { files: [new File([source], 'second.docx')] } });
  });
  await waitFor(() => expect(workers).toHaveLength(2));
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  expect(onError.mock.calls).toEqual([[new Error('Document was replaced'), { recoverable: true }]]);
});

test('a property-dialog command failure reports a recoverable error and keeps the document open', async () => {
  const ref = createRef<DocxEditorRef>();
  const failure = new Error('Watermark properties failed');
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  const view = render(<DocxEditor {...peerlessProps} ref={ref} onError={onError} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const read = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentRead').mockRejectedValueOnce(failure);
  restorers.push(() => read.mockRestore());
  await act(async () => { expect(await ref.current!.commands.execute('watermark', null)).toMatchObject({ ok: false }); });
  expect(onError.mock.calls).toEqual([[failure, { recoverable: true }]]);
  expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true');
});

test('a find-dialog search failure reports a recoverable error and keeps the document open', async () => {
  const ref = createRef<DocxEditorRef>();
  const failure = new Error('Search failed');
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  const view = render(<DocxEditor {...peerlessProps} ref={ref} onError={onError} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { await ref.current!.commands.execute('find', null); });
  const read = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'documentReadAt').mockRejectedValueOnce(failure);
  restorers.push(() => read.mockRestore());
  fireEvent.change(view.getByTestId('find-replace-dialog').querySelector('#find-text')!, { target: { value: 'cat' } });
  await waitFor(() => expect(onError.mock.calls).toEqual([[failure, { recoverable: true }]]));
  expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true');
});

test('worker save downloads the worker bytes with the host document name', async () => {
  const ref = createRef<DocxEditorRef>();
  const downloads: string[] = [];
  const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) { downloads.push(this.download); });
  restorers.push(() => click.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} ref={ref} documentName="draft.docx" />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  await act(async () => { expect(await ref.current!.commands.execute('save', null)).toMatchObject({ ok: true }); });
  expect(downloads).toEqual(['draft.docx']);
  expect(requests.filter(({ type }) => type === 'save')).toHaveLength(1);
});

test('worker print builds all pages and invokes the host after complete rendering', async () => {
  const ref = createRef<DocxEditorRef>();
  const onPrint = mock(() => {});
  const popup = { document: document.implementation.createHTMLDocument(), opener: null, closed: false,
    focus() {}, print: mock(() => {}), close: mock(() => {}) };
  const open = spyOn(window, 'open').mockReturnValue(popup as unknown as Window);
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({ drawImage() {} } as unknown as CanvasRenderingContext2D);
  const png = spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('data:image/png;base64,cGFnZQ==');
  const build = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'buildPages').mockImplementation(async (indices) => {
    surface.mock.calls.at(-1)![0].sessionOptions.prepareGeometry({ ...commandFrame(2), payload: {
      ...commandFrame(2).payload, stagedPresentation: { frameEpoch: 2, pages: indices.map((pageIndex) => ({
        page: { pageId: String(pageIndex + 1), pageIndex }, frameEpoch: 2,
        bitmap: { width: 200, height: 100, close() {} } as ImageBitmap,
      })) },
    } });
    return null as never;
  });
  restorers.push(() => open.mockRestore(), () => context.mockRestore(), () => png.mockRestore(), () => build.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} ref={ref} onPrint={onPrint} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  binding!.viewStore.getSnapshot().geometry = { identity: { session: 'print', incarnation: 1, epoch: 1 }, coversThroughSeq: 0,
    frameEpoch: 1, layoutRevision: 1, caret: null, selectionRects: [], pages: [0, 1, 2].map((pageIndex) => ({
      page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 100, width: 200, height: 100 },
      transform: [1, 0, 0, 1, 0, 0],
    })) };
  await act(async () => { expect(await ref.current!.commands.execute('print', null)).toMatchObject({ ok: true }); });
  expect(build.mock.calls.map(([indices]) => indices)).toEqual([[0], [1], [2]]);
  expect(popup.document.querySelectorAll('img')).toHaveLength(3);
  expect(popup.print).toHaveBeenCalledTimes(1);
  expect(onPrint).toHaveBeenCalledTimes(1);
});

test('overlapping prints during a provisional prefix both capture the whole document', async () => {
  const ref = createRef<DocxEditorRef>();
  const onError = mock(() => {});
  const popups = [0, 1].map(() => ({ document: document.implementation.createHTMLDocument(), opener: null, closed: false,
    focus() {}, print: mock(() => {}), close: mock(() => {}) }));
  let opened = 0;
  const open = spyOn(window, 'open').mockImplementation(() => popups[opened++] as unknown as Window);
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({ drawImage() {} } as unknown as CanvasRenderingContext2D);
  const png = spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('data:image/png;base64,cGFnZQ==');
  let release!: () => void;
  const completing = new Promise<void>((resolve) => { release = resolve; });
  const complete = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'completeLayout').mockImplementation(async () => {
    await completing;
    return null as never;
  });
  const geometry = (provisional: boolean): PeerlessFrameGeometry => ({
    identity: { session: 'print', incarnation: 1, epoch: 1 }, coversThroughSeq: 0,
    frameEpoch: provisional ? 1 : 2, layoutRevision: 1, layoutProvisional: provisional, caret: null, selectionRects: [],
    pages: Array.from({ length: provisional ? 3 : 5 }, (_, pageIndex) => ({
      page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 100, width: 200, height: 100 },
      transform: [1, 0, 0, 1, 0, 0],
    })),
  });
  const build = spyOn(yrs.ResidentEngineWorkerClient.prototype, 'buildPages').mockImplementation(async (indices) => {
    binding!.viewStore.getSnapshot().geometry = geometry(false);
    surface.mock.calls.at(-1)![0].sessionOptions.prepareGeometry({ ...commandFrame(2), payload: {
      ...commandFrame(2).payload, stagedPresentation: { frameEpoch: 2, pages: indices.map((pageIndex) => ({
        page: { pageId: String(pageIndex + 1), pageIndex }, frameEpoch: 2,
        bitmap: { width: 200, height: 100, close() {} } as ImageBitmap,
      })) },
    } });
    return null as never;
  });
  restorers.push(() => open.mockRestore(), () => context.mockRestore(), () => png.mockRestore(),
    () => complete.mockRestore(), () => build.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} ref={ref} onError={onError} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  binding!.viewStore.getSnapshot().geometry = geometry(true);
  await act(async () => {
    const first = ref.current!.commands.execute('print', null);
    const second = ref.current!.commands.execute('print', null);
    await waitFor(() => expect(complete).toHaveBeenCalledTimes(2));
    release();
    const outcomes = await Promise.all([first, second]);
    expect(outcomes).toMatchObject([{ ok: true }, { ok: true }]);
  });
  expect(onError).not.toHaveBeenCalled();
  for (const popup of popups) {
    expect(popup.document.querySelectorAll('img')).toHaveLength(5);
    expect(popup.print).toHaveBeenCalledTimes(1);
  }
});

test('worker context menu reads selection and table context without a session', async () => {
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const snapshot = binding!.viewStore.getSnapshot();
  snapshot.workerSelection = { anchor: { story: 'body', paraId: 'a', offset: 0 }, focus: { story: 'body', paraId: 'a', offset: 3 } };
  snapshot.editState = { ...snapshot.editState!, tableContext: { story: 'body', tableIndex: 0, rowIndex: 0,
    columnIndex: 0, rowCount: 2, columnCount: 2, hasMultiCellSelection: false, canSplitCell: true } };
  fireEvent.contextMenu(view.getByTestId('peerless-paged-editor'), { clientX: 40, clientY: 60 });
  expect(view.getByRole('menu')).toBeTruthy();
  expect(view.getByRole('menuitem', { name: /^Copy/ }).getAttribute('aria-disabled')).not.toBe('true');
  expect(view.getByRole('menuitem', { name: /Insert row above/ }).getAttribute('aria-disabled')).not.toBe('true');
  expect(coreSession).not.toHaveBeenCalled();
});

test('worker outline queries headings and navigates through worker selection commands', async () => {
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} showOutline />);
  await waitFor(() => expect(view.getByText('Worker heading')).toBeTruthy());
  fireEvent.click(view.getByText('Worker heading'));
  await waitFor(() => expect(requests.some((request) => request.type === 'edit' && request.command.kind === 'select')).toBe(true));
  expect(coreSession).not.toHaveBeenCalled();
});

test('hybrid compatibility opt-in preserves Document callbacks and warns once', () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => { warn.mockRestore(); resetEngineChoiceForTests(); });
  const first = render(<DocxEditor {...peerlessProps} {...{ [LEGACY_SYNC_API_PROP]: true }} onChange={() => {}} />);
  expect(coreSession).toHaveBeenCalled();
  expect(workers).toHaveLength(0);
  first.unmount();
  render(<DocxEditor {...peerlessProps} {...{ [LEGACY_SYNC_API_PROP]: true }} onChange={() => {}} />);
  expect(warn.mock.calls.filter(([message]) => String(message).includes(LEGACY_SYNC_API_PROP))).toHaveLength(1);
});

test('collaboration uses the hybrid and warns once across two mounts', () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => warn.mockRestore());
  const first = render(<DocxEditor {...peerlessProps} collaboration={{ onReplica() {} }} />);
  expect(coreSession).toHaveBeenCalled();
  expect(workers).toHaveLength(0);
  first.unmount();
  render(<DocxEditor {...peerlessProps} collaboration={{ onReplica() {} }} />);
  expect(warn.mock.calls.filter(([message]) => String(message).includes('worker-only collaboration'))).toHaveLength(1);
});

test('peerless collaboration uses the hybrid without the explicit opt-in', () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => warn.mockRestore());
  const onError = mock(() => {});
  const view = render(<DocxEditor {...peerlessProps} collaboration={{ onReplica() {} }} onError={onError} />);
  expect(view.queryByRole('alert')).toBeNull();
  expect(onError).not.toHaveBeenCalled();
  expect(coreSession).toHaveBeenCalled();
  expect(workers).toHaveLength(0);
  expect(warn.mock.calls).toEqual([[expect.stringContaining('worker-only collaboration')]]);
});

test('collaboration presence remounts between default and hybrid routes', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  view.rerender(<DocxEditor {...peerlessProps} collaboration={{ onReplica() {} }} />);
  await waitFor(() => expect(workers[0]!.terminate).toHaveBeenCalled());
  expect(view.queryByTestId('peerless-paged-editor')).toBeNull();
  expect(coreSession).toHaveBeenCalled();
  coreSession.mockClear();
  view.rerender(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(workers).toHaveLength(2));
  expect(coreSession).not.toHaveBeenCalled();
});

test('in-thread opt-in warns once across two mounts and accepts Document callbacks', () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => warn.mockRestore());
  const first = render(<DocxEditor {...peerlessProps} experimentalWorkerOpen={false} onChange={() => {}} />);
  expect(coreSession).toHaveBeenCalled();
  expect(workers).toHaveLength(0);
  first.unmount();
  render(<DocxEditor {...peerlessProps} experimentalWorkerOpen={false} onChange={() => {}} />);
  expect(warn.mock.calls.filter(([message]) => String(message).includes('experimentalWorkerOpen={false}'))).toHaveLength(1);
});

test('in-thread opt-in remounts a mounted default editor and can return to the worker', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  view.rerender(<DocxEditor {...peerlessProps} experimentalWorkerOpen={false} onChange={() => {}} />);
  await waitFor(() => expect(workers[0]!.terminate).toHaveBeenCalled());
  expect(coreSession).toHaveBeenCalled();
  expect(view.queryByTestId('peerless-paged-editor')).toBeNull();
  coreSession.mockClear();
  view.rerender(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(workers).toHaveLength(2));
  expect(coreSession).not.toHaveBeenCalled();
});

test('capability fallback accepts sync callbacks and emits only its capability warning', () => {
  setMissingWorkerCapabilitiesForTests(['OffscreenCanvas']);
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const onError = mock(() => {});
  restorers.push(() => warn.mockRestore());
  const invalid = { onChange() {} } as unknown as DocxEditorProps;
  const view = render(<DocxEditor {...peerlessProps} {...invalid} onError={onError} />);
  expect(view.queryByRole('alert')).toBeNull();
  expect(coreSession).toHaveBeenCalled();
  expect(onError).not.toHaveBeenCalled();
  expect(warn.mock.calls).toEqual([[expect.stringContaining('browser features are unavailable: OffscreenCanvas')]]);
});

test('hybrid opt-in takes precedence over collaboration and sync callbacks', () => {
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => warn.mockRestore());
  render(<DocxEditor {...peerlessProps} {...{ [LEGACY_SYNC_API_PROP]: true }} collaboration={{}} onChange={() => {}} />);
  expect(coreSession).toHaveBeenCalled();
  expect(workers).toHaveLength(0);
  expect(warn.mock.calls).toEqual([[expect.stringContaining(`${LEGACY_SYNC_API_PROP}={true}`)]]);
});

test('unrelated prop changes keep the route until a routing input changes', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  setMissingWorkerCapabilitiesForTests(['OffscreenCanvas']);
  view.rerender(<DocxEditor {...peerlessProps} documentName="Changed" onDocumentChange={() => {}} />);
  expect(workers).toHaveLength(1);
  expect(coreSession).not.toHaveBeenCalled();
  setMissingWorkerCapabilitiesForTests([]);
  const invalid = { onChange() {} } as unknown as DocxEditorProps;
  view.rerender(<DocxEditor {...peerlessProps} {...invalid} />);
  expect(view.getByRole('alert').textContent).toContain(LEGACY_SYNC_API_PROP);
  await waitFor(() => expect(workers[0]!.terminate).toHaveBeenCalled());
  view.rerender(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(workers).toHaveLength(2));
});

test.each(['Worker', 'OffscreenCanvas'] as const)('peerless mount falls back with the existing warning when %s is missing', (capability) => {
  setMissingWorkerCapabilitiesForTests([capability]);
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  restorers.push(() => { warn.mockRestore(); resetEngineChoiceForTests(); });
  const view = render(<DocxEditor {...peerlessProps} />);
  expect(view.queryByTestId('peerless-paged-editor')).toBeNull();
  expect(coreSession).toHaveBeenCalled();
  expect(workers).toHaveLength(0);
  expect(warn.mock.calls.some(([message]) => String(message).includes(`browser features are unavailable: ${capability}`))).toBe(true);
});



test('peerless shell catches chrome failures and reports worker errors in the banner', async () => {
  const onError = mock(() => {});
  function BrokenChrome(): never { throw new Error('chrome failed'); }
  const log = spyOn(console, 'error').mockImplementation(() => {});
  restorers.push(() => log.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} toolbar={<BrokenChrome />} onError={onError} />);
  await waitFor(() => expect(onError).toHaveBeenCalled());
  expect(view.getByRole('heading', { name: 'Something went wrong' })).toBeTruthy();
  view.unmount();
  const next = render(<DocxEditor {...peerlessProps} onError={onError} />);
  await waitFor(() => expect(next.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  act(() => { surface.mock.calls.at(-1)![0].sessionOptions.onError?.(new Error('worker failed to open')); });
  await waitFor(() => expect(next.getByText('worker failed to open')).toBeTruthy());
});

test('a recoverable presentation failure reaches the host without closing the document', async () => {
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  const view = render(<DocxEditor {...peerlessProps} onError={onError} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const stall = new Error('Full frame responses did not satisfy requested coverage and revision');
  act(() => { surface.mock.calls.at(-1)![0].sessionOptions.onRecoverableError!(stall); });
  expect(onError.mock.calls).toEqual([[stall, { recoverable: true }]]);
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true');
  expect(view.queryByText(stall.message)).toBeNull();
});

for (const [name, update, message] of [
  ['read-only', { readOnly: true }, 'DOCX read-only transition was rejected'],
  ['render', { showHiddenText: true }, 'DOCX render environment was rejected: stale-version'],
  ['font', { measurementFontProvider: { resolve: () => undefined } }, 'DOCX font environment was rejected: stale-version'],
] as const) {
  test(`a rejected ${name} environment reaches the host as recoverable`, async () => {
    const onError = mock(() => {});
    const view = render(<DocxEditor {...peerlessProps} onError={onError} />);
    await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
    workers[0].refuseNextEdit = true;
    view.rerender(<DocxEditor {...peerlessProps} {...update} onError={onError} />);
    await waitFor(() => expect(onError).toHaveBeenCalledWith(expect.objectContaining({ message }), { recoverable: true }));
    expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true');
    expect(view.queryByText(message)).toBeNull();
  });
}

test('a failed open reaches the host without recoverable metadata', async () => {
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  const view = render(<DocxEditor {...peerlessProps} onError={onError} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  const failure = new Error('worker failed to open');
  act(() => { surface.mock.calls.at(-1)![0].sessionOptions.onError!(failure); });
  expect(onError.mock.calls).toEqual([[failure]]);
  await waitFor(() => expect(view.getByText(failure.message)).toBeTruthy());
  expect(view.queryByTestId('peerless-paged-editor')).toBeNull();
});

test('a failed Blob read ends opening and reaches the shell error banner', async () => {
  const conversion = spyOn(utils, 'toArrayBuffer').mockRejectedValue(new Error('source read failed'));
  restorers.push(() => conversion.mockRestore());
  const onError = mock(() => {});
  const view = render(<DocxEditor {...peerlessProps} documentBuffer={new Blob(['source'])} onError={onError} />);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  await waitFor(() => expect(view.getByText('source read failed')).toBeTruthy());
  expect(view.queryByTestId('canvas-renderer-loading')).toBeNull();
  expect(workers).toHaveLength(0);
});

test('d8a_peerless_imports_use_exported_docx_entry_points', async () => {
  const files = [
    './PeerlessDocxEditor.tsx', './PeerlessDocxEditor.test.tsx', '../../DocxEditor.tsx',
    '../hooks/useRustMeasurement.ts', '../../../index.ts',
    '../../../../../docx/src/yrs/index.ts', '../../../../../docx/src/yrs/yrsSessionFacade.ts',
    '../../../../../docx/src/editor/computeLayout.ts', '../../../../../../apps/desktop/src/Editor.tsx',
  ];
  for (const file of files) {
    const code = await Bun.file(new URL(file, import.meta.url)).text();
    for (const match of code.matchAll(/from ['"](@betteroffice\/docx(?:\/[^'"]*)?)['"]/g)) {
      const suffix = match[1]!.slice('@betteroffice/docx'.length);
      expect(Object.hasOwn(docxPackage.exports, suffix ? `.${suffix}` : '.')).toBe(true);
    }
  }
  expect(typeof peerless.createResidentCommandConnection).toBe('function');
  expect(typeof peerless.createDocxWorkerEditSession).toBe('function');
});

test('d8a_command_init_uses_opened_sections_and_fonts', async () => {
  const metadata = hostMetadata();
  const first = { pageWidth: 9000, pageHeight: 14000, marginLeft: 720 };
  const second = { pageWidth: 16000, pageHeight: 10000, orientation: 'landscape', marginTop: 360 };
  metadata.envelope.document.package.document.sections = [
    { id: 'first', contentStart: 0, contentEnd: 0, properties: first },
    { id: 'second', contentStart: 0, contentEnd: 0, properties: second },
  ];
  metadata.referencedFonts = ['Embedded Serif', 'Authored Sans'];
  hostJson = JSON.stringify(metadata);
  fontRequirements = [
    { key: 'Embedded Serif|0|0', family: 'Embedded Serif', bold: false, italic: false },
    { key: 'Authored Sans|1|0', family: 'Authored Sans', bold: true, italic: false },
  ];
  const embedded = Uint8Array.of(11, 12).buffer;
  const bundled = Uint8Array.of(21, 22).buffer;
  const faces = spyOn(utils, 'extractEmbeddedFontFaces').mockResolvedValue([
    { family: 'Embedded Serif', weight: 'normal', style: 'normal', data: embedded, subsetted: false },
  ]);
  const resolve = mock((family: string, bold: boolean, italic: boolean) =>
    family === 'Authored Sans' && bold && !italic ? async () => bundled : undefined);
  const main = spyOn(yrs, 'createYrsSession').mockImplementation(() => {
    throw new Error('main document created');
  });
  restorers.push(() => faces.mockRestore(), () => main.mockRestore());
  render(<DocxEditor {...peerlessProps} measurementFontProvider={{ resolve }} />);
  await waitFor(() => expect(binding).toBeDefined());
  const init = requests.find((request) => request.type === 'commandInit');
  if (!init || init.type !== 'commandInit') throw new Error('command initialization missing');
  const layout = JSON.parse(init.request.layoutInput);
  expect(layout.regions.sections).toEqual([
    { sectionId: 'first', properties: first },
    { sectionId: 'second', properties: second },
    { sectionId: 'second', properties: second },
  ]);
  expect(init.request.fonts.filter((font) => font instanceof Uint8Array)).toEqual(
    expect.arrayContaining([new Uint8Array(embedded), new Uint8Array(bundled)]));
  expect(Object.keys(layout.measurement.fontChains).sort()).toEqual(fontRequirements.map(({ key }) => key).sort());
  expect(Object.values(layout.measurement.fontChains).sort()).toEqual([[0], [1]]);
  expect(resolve).toHaveBeenCalledWith('Authored Sans', true, false);
  const requirements = requests.find((request) => request.type === 'fontRequirements');
  if (!requirements || requirements.type !== 'fontRequirements') throw new Error('font requirements missing');
  expect(JSON.parse(requirements.layoutInput).regions).toEqual(layout.regions);
  expect(main).not.toHaveBeenCalled();
});

test('d8a_readonly_retries_after_refused_predecessor', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const current = binding!;
  const worker = workers[0];
  worker.holdEdits = true;
  worker.refuseNextEdit = true;
  const edit = current.commands.submit({ kind: 'insertText', payload: { text: 'refused' } });
  if (edit.status !== 'submitted') throw new Error('edit not submitted');
  await waitFor(() => expect(worker.pendingReplies).toHaveLength(1));
  view.rerender(<DocxEditor {...peerlessProps} readOnly />);
  await waitFor(() => expect(worker.pendingReplies).toHaveLength(2));
  expect(current.session.state.editState?.readOnly).toBe(false);
  worker.holdEdits = false;
  for (const response of worker.pendingReplies.splice(0)) {
    worker.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  expect((await edit.ticket.ack).status).toBe('refused');
  await waitFor(() => expect(current.session.state.editState?.readOnly).toBe(true));
  const commands = requests.filter((request): request is Extract<ResidentEngineWorkerRequest, { type: 'edit' }> => request.type === 'edit');
  expect(commands.map(({ command }) => command.kind)).toEqual(['insertText', 'setReadOnly']);
  expect(commands[1].command.afterSeq).toBeUndefined();
  expect(worker.text).toBe('');
  expect(workers).toHaveLength(1);
  view.rerender(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(current.session.state.editState?.readOnly).toBe(false));
  view.rerender(<DocxEditor {...peerlessProps} readOnly />);
  await waitFor(() => expect(current.session.state.editState?.readOnly).toBe(true));
});

test('d8a_readonly_unlock_precedes_typing_with_lock_ack_withheld', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const current = binding!;
  const worker = workers[0];
  worker.holdEdits = true;
  view.rerender(<DocxEditor {...peerlessProps} readOnly />);
  await waitFor(() => expect(worker.pendingReplies).toHaveLength(1));
  view.rerender(<DocxEditor {...peerlessProps} />);
  const edit = current.commands.submit({ kind: 'insertText', payload: { text: 'after unlock' } });
  if (edit.status !== 'submitted') throw new Error('edit not submitted');
  await waitFor(() => expect(worker.pendingReplies).toHaveLength(3));
  const commands = requests.filter((request): request is Extract<ResidentEngineWorkerRequest, { type: 'edit' }> => request.type === 'edit');
  expect(commands.map(({ command }) => command.kind)).toEqual(['setReadOnly', 'setReadOnly', 'insertText']);
  expect(commands.slice(0, 2).map(({ command }) => command.payload)).toEqual([
    { readOnly: true }, { readOnly: false },
  ]);
  expect(worker.text).toBe('after unlock');
  worker.holdEdits = false;
  for (const response of worker.pendingReplies.splice(0)) {
    worker.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  expect((await edit.ticket.ack).status).toBe('applied');
  await waitFor(() => expect(current.session.state.editState?.readOnly).toBe(false));
  expect(requests.filter(({ type }) => type === 'edit')).toHaveLength(3);
});

test('d8a_readonly_unlock_precedes_parent_layout_input', async () => {
  const submitted: { edit?: ReturnType<PeerlessEditSessionBinding['commands']['submit']> } = {};
  let passiveReadOnly = true;
  function Parent({ readOnly, input = false }: { readOnly: boolean; input?: boolean }) {
    useLayoutEffect(() => {
      if (!input) return;
      expect(passiveReadOnly).toBe(true);
      submitted.edit = binding!.commands.submit({ kind: 'insertText', payload: { text: 'after unlock' } });
    }, [readOnly, input]);
    useEffect(() => { passiveReadOnly = readOnly; }, [readOnly]);
    return <DocxEditor {...peerlessProps} readOnly={readOnly} />;
  }
  const view = render(<Parent readOnly />);
  await waitFor(() => expect(binding).toBeDefined());
  expect(binding!.session.state.editState?.readOnly).toBe(true);
  view.rerender(<Parent readOnly={false} input />);
  const edit = submitted.edit;
  if (!edit || edit.status !== 'submitted') throw new Error('edit not submitted');
  expect((await edit.ticket.ack).status).toBe('applied');
  expect(workers[0].text).toBe('after unlock');
  const commands = requests.filter((request): request is Extract<ResidentEngineWorkerRequest, { type: 'edit' }> => request.type === 'edit');
  expect(commands.map(({ command }) => command.kind)).toEqual(['setReadOnly', 'insertText']);
  expect(commands[0].command.payload).toEqual({ readOnly: false });
});

test('d8a_readonly_lock_precedes_parent_layout_input', async () => {
  const submitted: { edit?: ReturnType<PeerlessEditSessionBinding['commands']['submit']> } = {};
  let passiveReadOnly = false;
  function Parent({ readOnly, input = false }: { readOnly: boolean; input?: boolean }) {
    useLayoutEffect(() => {
      if (!input) return;
      expect(passiveReadOnly).toBe(false);
      submitted.edit = binding!.commands.submit({ kind: 'insertText', payload: { text: 'after lock' } });
    }, [readOnly, input]);
    useEffect(() => { passiveReadOnly = readOnly; }, [readOnly]);
    return <DocxEditor {...peerlessProps} readOnly={readOnly} />;
  }
  const view = render(<Parent readOnly={false} />);
  await waitFor(() => expect(binding).toBeDefined());
  expect(binding!.session.state.editState?.readOnly).toBe(false);
  view.rerender(<Parent readOnly input />);
  const edit = submitted.edit;
  expect(edit).toEqual({ status: 'ignored', reason: 'read-only' });
  expect(workers[0].text).toBe('');
  await waitFor(() => expect(requests.filter(({ type }) => type === 'edit')).toHaveLength(1));
  const commands = requests.filter((request): request is Extract<ResidentEngineWorkerRequest, { type: 'edit' }> => request.type === 'edit');
  expect(commands.map(({ command }) => command.kind)).toEqual(['setReadOnly']);
  expect(commands[0].command.payload).toEqual({ readOnly: true });
});

test('d8a_edits_survive_readonly_and_budget_changes', async () => {
  const view = render(<DocxEditor {...peerlessProps} memoryBudget={{ workerLimitBytes: 1024 }} />);
  await waitFor(() => expect(binding).toBeDefined());
  const current = binding!;
  workers[0].holdEdits = true;
  const edit = current.commands.submit({ kind: 'insertText', payload: { text: 'live edit' } });
  expect(edit.status).toBe('submitted');
  if (edit.status !== 'submitted') throw new Error('edit not submitted');
  await waitFor(() => expect(workers[0].text).toBe('live edit'));
  const options = surface.mock.calls[0][0].sessionOptions;
  view.rerender(<DocxEditor {...peerlessProps} readOnly memoryBudget={{ workerLimitBytes: 2048 }} />);
  expect(surface.mock.calls.at(-1)![0].sessionOptions).toBe(options);
  expect(workers).toHaveLength(1);
  expect(requests.filter(({ type }) => type === 'openSource')).toHaveLength(1);
  expect(requests.filter(({ type }) => type === 'commandInit')).toHaveLength(1);
  expect(workers[0].terminate).not.toHaveBeenCalled();
  workers[0].holdEdits = false;
  for (const response of workers[0].pendingReplies.splice(0)) {
    workers[0].onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  await edit.ticket.ack;
  await waitFor(() => expect(current.session.state.editState?.readOnly).toBe(true));
  expect(binding).toBe(current);
  expect(workers[0].text).toBe('live edit');
  const commands = requests.filter((request): request is Extract<ResidentEngineWorkerRequest, { type: 'edit' }> => request.type === 'edit');
  expect(commands.map(({ command }) => command.kind)).toEqual(['insertText', 'setReadOnly']);
  expect(commands[1].command.afterSeq).toBeUndefined();
  view.rerender(<DocxEditor {...peerlessProps} memoryBudget={{ workerLimitBytes: 4096 }} />);
  await waitFor(() => expect(current.session.state.editState?.readOnly).toBe(false));
  expect(workers[0].text).toBe('live edit');
  expect(workers).toHaveLength(1);
  view.unmount();
  await waitFor(() => expect(workers[0].terminate).toHaveBeenCalledTimes(1));
});

test('d8a_canvas_lookup_uses_page_identity_without_scanning', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const host = view.container.firstElementChild!;
  const scan = spyOn(host, 'querySelectorAll').mockImplementation(() => {
    throw new Error('canvas scan');
  });
  restorers.push(() => scan.mockRestore());
  const options = surface.mock.calls[0][0].sessionOptions;
  expect(options.canvasForPage({ pageId: '17', pageIndex: 0 })).toBe(view.getByTestId('page-canvas') as HTMLCanvasElement);
  expect(options.canvasForPage({ pageId: '18', pageIndex: 0 })).toBeNull();
  expect(options.canvasForPage({ pageId: '17"]', pageIndex: 0 })).toBeNull();
  expect(scan).not.toHaveBeenCalled();
});

test('d8a_rejected_presentation_can_retry_prepared_frame', async () => {
  render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const options = surface.mock.calls[0][0].sessionOptions;
  expect(options.prepareGeometry(commandFrame(1))?.frameEpoch).toBe(1);
  const frame = commandFrame(3, 2);
  frame.payload.pageFrames = [commandFrame(2).payload.frame];
  const state: { ready: boolean; installed?: PeerlessFrameGeometry } = { ready: false };
  const present = (value: DocxCommandFrame) => {
    const geometry = options.prepareGeometry(value);
    if (!geometry || !state.ready) return false;
    state.installed = geometry;
    return true;
  };
  expect(present(frame)).toBe(false);
  expect(state.installed).toBeUndefined();
  state.ready = true;
  expect(present(frame)).toBe(true);
  expect(state.installed?.frameEpoch).toBe(3);
  expect(state.installed?.coversThroughSeq).toBe(2);
  expect(options.prepareGeometry(frame)).toBe(state.installed);
  expect(options.prepareGeometry(commandFrame(4, 3))?.frameEpoch).toBe(4);
});

test('prepared geometry keeps worker selection rectangles for the page overlay', async () => {
  render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const frame = commandFrame(1);
  Object.assign(frame.payload, { selectionRects: [{ pageIndex: 0, x: 12, y: 20, width: 30, height: 16 }] });
  const options = surface.mock.calls[0][0].sessionOptions;
  expect(options.prepareGeometry(frame)?.selectionRects).toEqual([{ x: 12, y: 20, width: 30, height: 16 }]);
});

test('a typing presentation decodes each buffer once across validation, geometry and echo', async () => {
  render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const options = surface.mock.calls[0][0].sessionOptions;
  const painter = createEchoPainter({ canvasForPage: options.canvasForPage });
  const presentation = createFramePresentation({ viewStore: binding!.viewStore, echoPainter: painter,
    canvasForPage: options.canvasForPage, prepareGeometry: options.prepareGeometry,
    requestFullFrame() { throw new Error('Unexpected full frame request'); } });
  const decode = spyOn(frameRender, 'decodeFrameDelta');
  try {
    const initial = commandFrame(1);
    initial.payload.stagedPresentation = { frameEpoch: 1, pages: [] };
    expect(presentation.present(initial, [], [])).toBe(true);
    decode.mockClear();
    const typed = commandFrame(4, 1);
    const buffers = [commandFrame(2, 1).payload.frame, commandFrame(3, 2).payload.frame, typed.payload.frame];
    const data = new DataView(typed.payload.frame);
    data.setBigUint64(40, 3n, true);
    typed.payload.pageFrames = buffers.slice(0, 2);
    typed.payload.stagedPresentation = { frameEpoch: 4, pages: [] };
    expect(presentation.present(typed, [], [])).toBe(true);
    expect(buffers.map((bytes) => decode.mock.calls.filter(([input]) => input === bytes).length)).toEqual([1, 1, 1]);
    decode.mockClear();
    data.setBigUint64(32, 5n, true);
    data.setBigUint64(40, 4n, true);
    const next = commandFrame(5, 4);
    next.payload.frame = typed.payload.frame;
    next.payload.stagedPresentation = { frameEpoch: 5, pages: [] };
    expect(presentation.present(next, [], [])).toBe(true);
    expect(decode).toHaveBeenCalledTimes(1);
    expect(decode.mock.calls[0]![0]).toBe(next.payload.frame);
  } finally {
    decode.mockRestore();
    presentation.dispose();
    painter.dispose();
  }
});

test('hybrid opt-in remounts the editor when toggled', async () => {
  const renderer = spyOn(display, 'useCanvasRenderer');
  restorers.push(() => renderer.mockRestore());
  const view = render(<DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }} showToolbar={false} />);
  expect(core.useYrsCoreSession).toHaveBeenCalled();
  expect(renderer).toHaveBeenCalled();
  expect(paged.PeerlessPagedEditor).not.toHaveBeenCalled();
  expect(workers).toHaveLength(0);
  coreSession.mockClear();
  renderer.mockClear();
  view.rerender(<DocxEditor {...peerlessProps} showToolbar={false} />);
  await waitFor(() => expect(paged.PeerlessPagedEditor).toHaveBeenCalled());
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  expect(core.useYrsCoreSession).not.toHaveBeenCalled();
  expect(renderer).not.toHaveBeenCalled();
  view.rerender(<DocxEditor {...{ [LEGACY_SYNC_API_PROP]: true }} showToolbar={false} />);
  expect(core.useYrsCoreSession).toHaveBeenCalled();
  expect(renderer).toHaveBeenCalled();
  expect(view.queryByTestId('peerless-paged-editor')).toBeNull();
  expect(workers).toHaveLength(1);
});

test('context paste refuses a target changed while clipboard permission is pending', async () => {
  const onError = mock(() => {});
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} onError={onError} />);
  await waitFor(() => expect(binding).toBeDefined());
  const snapshot = binding!.viewStore.getSnapshot();
  snapshot.workerSelection = { anchor: { story: 'body', paraId: 'a', offset: 0 }, focus: { story: 'body', paraId: 'a', offset: 3 } };
  let finish!: (text: string) => void;
  const read = spyOn(navigator.clipboard, 'readText').mockImplementation(() => new Promise<string>((resolve) => { finish = resolve; }));
  restorers.push(() => read.mockRestore());
  fireEvent.contextMenu(view.getByTestId('peerless-paged-editor'));
  fireEvent.click(view.getByRole('menuitem', { name: /^Paste/ }));
  await waitFor(() => expect(read).toHaveBeenCalled());
  snapshot.version += 1;
  await act(async () => finish('clipboard text'));
  expect(requests.filter((request) => request.type === 'edit')).toHaveLength(0);
  expect(onError).toHaveBeenCalledWith(expect.objectContaining({ message: 'Context menu target changed' }), { recoverable: true });
  expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true');
});

test('worker shell reserves the widest scaled page and retains host styling', async () => {
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} initialZoom={1.5} />);
  await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
  binding!.viewStore.getSnapshot().geometry = { identity: { session: 'sizes', incarnation: 1, epoch: 1 }, coversThroughSeq: 0,
    frameEpoch: 1, layoutRevision: 1, caret: null, selectionRects: [], pages: [816, 1200].map((width, pageIndex) => ({
      page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 100, width, height: 100 },
      transform: [1, 0, 0, 1, 0, 0],
    })) };
  view.rerender(<DocxEditor {...peerlessProps} showToolbar={false} initialZoom={1.5} showOutline
    className="host-editor" style={{ height: 300, width: 400 }} pluginSidebarItems={[{ id: 'card', anchorPos: 0, hidden: false, render: () => null }]} />);
  await waitFor(() => {
    const root = view.getByTestId('docx-editor');
    expect(root.classList.contains('host-editor')).toBe(true);
    expect(root.style.height).toBe('300px');
    expect(root.style.backgroundColor).toBe('var(--doc-bg)');
    const layout = view.container.querySelector('.docx-editor__scroll-container')!.firstElementChild as HTMLElement;
    expect(layout.style.minWidth).toBe(`${1200 * 1.5 + 2 * OUTLINE_RESERVED_SPACE + 2 * SIDEBAR_DOCUMENT_SHIFT}px`);
  });
  expect(coreSession).not.toHaveBeenCalled();
});


test('suspended_readonly_render_preserves_committed_editability', async () => {
  const pending = new Promise<void>(() => {});
  let requestLock!: () => void;
  let renders = 0;
  function Suspend({ active }: { active: boolean }) {
    if (active) { renders += 1; throw pending; }
    return null;
  }
  function Parent() {
    const [locked, setLocked] = useState(false);
    requestLock = () => startTransition(() => setLocked(true));
    return <Suspense fallback={<span>pending</span>}>
      <DocxEditor {...peerlessProps} readOnly={locked} />
      <Suspend active={locked} />
    </Suspense>;
  }
  const view = render(<Parent />);
  await waitFor(() => expect(binding).toBeDefined());
  await act(async () => { requestLock(); });
  await waitFor(() => expect(renders).toBeGreaterThan(0));
  const edit = binding!.commands.submit({ kind: 'insertText', payload: { text: 'committed' } });
  if (edit.status !== 'submitted') throw new Error('edit not submitted');
  expect((await edit.ticket.ack).status).toBe('applied');
  expect(workers[0].text).toBe('committed');
  expect(requests.filter((request) => request.type === 'edit').map((request) => request.command.kind)).toEqual(['insertText']);
  view.unmount();
});


test('withheld_refused_edit_cannot_cancel_unlock_or_following_input', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  const current = binding!;
  const worker = workers[0];
  worker.holdEdits = true;
  view.rerender(<DocxEditor {...peerlessProps} readOnly />);
  const refused = current.session.submit('insertText', { text: 'locked' }, {
    after: current.admission.getSnapshot().tail,
  });
  await waitFor(() => expect(worker.pendingReplies).toHaveLength(2));
  view.rerender(<DocxEditor {...peerlessProps} />);
  const edit = current.commands.submit({ kind: 'insertText', payload: { text: 'valid' } });
  if (edit.status !== 'submitted') throw new Error('edit not submitted');
  await waitFor(() => expect(worker.pendingReplies).toHaveLength(4));
  expect(worker.text).toBe('valid');
  worker.holdEdits = false;
  for (const response of worker.pendingReplies.splice(0)) {
    worker.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  expect(await refused.ack).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect((await edit.ticket.ack).status).toBe('applied');
  expect(current.session.state.editState?.readOnly).toBe(false);
  expect(workers).toHaveLength(1);
});

test('rapid_mode_changes_keep_queued_input_on_its_committed_side', async () => {
  const view = render(<DocxEditor {...peerlessProps} readOnly />);
  await waitFor(() => expect(binding).toBeDefined());
  const current = binding!;
  const worker = workers[0];
  worker.holdEdits = true;
  view.rerender(<DocxEditor {...peerlessProps} />);
  const edit = current.commands.submit({ kind: 'insertText', payload: { text: 'valid' } });
  if (edit.status !== 'submitted') throw new Error('edit not submitted');
  view.rerender(<DocxEditor {...peerlessProps} readOnly />);
  expect(current.commands.submit({ kind: 'insertText', payload: { text: 'locked' } })).toEqual({ status: 'ignored', reason: 'read-only' });
  await waitFor(() => expect(worker.pendingReplies).toHaveLength(3));
  worker.holdEdits = false;
  for (const response of worker.pendingReplies.splice(0)) {
    worker.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  expect((await edit.ticket.ack).status).toBe('applied');
  expect(worker.text).toBe('valid');
  expect(current.session.state.editState?.readOnly).toBe(true);
});

test('suspended_mode_does_not_change_an_inflight_open', async () => {
  let release!: () => void;
  const prepared = new Promise<void>((resolve) => { release = resolve; });
  let preparing = false;
  const create = measurement.createRustMeasureSource;
  const gate = spyOn(measurement, 'createRustMeasureSource').mockImplementation((options) => {
    const provider = create(options);
    const prepare = provider.prepareFontRequirements.bind(provider);
    provider.prepareFontRequirements = async (requirements) => {
      preparing = true;
      await prepared;
      return prepare(requirements);
    };
    return provider;
  });
  restorers.push(() => gate.mockRestore());
  const pending = new Promise<void>(() => {});
  let requestLock!: () => void;
  let suspended = false;
  function Suspend({ active }: { active: boolean }) {
    if (active) { suspended = true; throw pending; }
    return null;
  }
  function Parent() {
    const [locked, setLocked] = useState(false);
    requestLock = () => startTransition(() => setLocked(true));
    return <Suspense fallback={<span>pending</span>}>
      <DocxEditor {...peerlessProps} readOnly={locked} />
      <Suspend active={locked} />
    </Suspense>;
  }
  render(<Parent />);
  await waitFor(() => expect(preparing).toBe(true));
  await act(async () => { requestLock(); });
  await waitFor(() => expect(suspended).toBe(true));
  release();
  await waitFor(() => expect(binding).toBeDefined());
  expect(binding!.session.state.editState?.readOnly).toBe(false);
  const edit = binding!.commands.submit({ kind: 'insertText', payload: { text: 'valid' } });
  if (edit.status !== 'submitted') throw new Error('edit not submitted');
  expect((await edit.ticket.ack).status).toBe('applied');
  expect(workers[0].text).toBe('valid');
});

test('toolbar and owned history shortcuts dispatch worker history', async () => {
  const view = render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  binding!.viewStore.getSnapshot().editState!.canUndo = true;
  view.rerender(<DocxEditor {...peerlessProps} documentName="History" />);
  fireEvent.click(view.getByRole('button', { name: /^Undo/ }));
  await waitFor(() => expect(requests.some((request) => request.type === 'edit' && request.command.kind === 'undo')).toBe(true));
  const input = document.createElement('textarea');
  input.className = 'paged-editor__peerless-input';
  view.getByTestId('peerless-paged-editor').appendChild(input);
  fireEvent.keyDown(input, { key: 'z', ctrlKey: true });
  await waitFor(() => expect(requests.filter((request) => request.type === 'edit' && request.command.kind === 'undo')).toHaveLength(2));
  expect(coreSession).not.toHaveBeenCalled();
});

test('shell find opens its dialog and exposes worker host search methods', async () => {
  const ref = createRef<DocxEditorRef>();
  const view = render(<DocxEditor {...peerlessProps} ref={ref} />);
  await waitFor(() => expect(binding).toBeDefined());
  await act(async () => { expect(await ref.current!.commands.execute('find', null)).toMatchObject({ ok: true, status: 'opened' }); });
  expect(view.getByRole('dialog')).toBeTruthy();
  const changed = mock(() => {});
  ref.current!.onSearchChange(changed);
  await act(async () => { expect(await ref.current!.search('cat')).toMatchObject({ total: 0, current: -1 }); });
  expect(changed).toHaveBeenCalled();
  expect(requests.some(request => request.type === 'documentRead' && request.read.kind === 'searchText')).toBe(true);
  ref.current!.clearSearch();
  expect(ref.current!.getSearchState()).toBeNull();
});

test('shell surface receives controller table navigation', async () => {
  render(<DocxEditor {...peerlessProps} />);
  await waitFor(() => expect(binding).toBeDefined());
  surface.mock.calls.at(-1)![0].onNavigateCell!('backward');
  await waitFor(() => expect(requests.some(request => request.type === 'edit' && request.command.kind === 'navigateCell')).toBe(true));
  const request = requests.find(request => request.type === 'edit' && request.command.kind === 'navigateCell');
  expect(request?.type === 'edit' && request.command.payload).toEqual({ direction: 'backward', extend: false, createLastRow: false });
});

test('viewing mode initializes worker read-only and internal mode changes use its barrier', async () => {
  const view = render(<DocxEditor {...peerlessProps} mode="viewing" />);
  await waitFor(() => expect(binding).toBeDefined());
  expect(workers[0].initial!.editState.readOnly).toBe(true);
  const prior = binding;
  view.unmount();
  const ref = createRef<DocxEditorRef>();
  render(<DocxEditor {...peerlessProps} ref={ref} />);
  await waitFor(() => expect(binding).not.toBe(prior));
  await act(async () => { expect(await ref.current!.commands.execute('editingMode', { mode: 'viewing' })).toMatchObject({ ok: true }); });
  await waitFor(() => expect(binding!.admission.getSnapshot().readOnly).toBe(true));
  expect(binding!.commands.submit({ kind: 'insertText', payload: { text: 'locked' } })).toEqual({ status: 'ignored', reason: 'read-only' });
  await act(async () => { expect(await ref.current!.commands.execute('editingMode', { mode: 'editing' })).toMatchObject({ ok: true }); });
  await waitFor(() => expect(binding!.admission.getSnapshot().readOnly).toBe(false));
});

test('viewing mode reports read-only to worker plugins', async () => {
  let context: DocxPluginContext<null> | undefined;
  const events: { mode: string; readOnly: boolean; snapshot: boolean }[] = [];
  const plugin = defineDocxPlugin({ id: 'mode-check', createState: () => null, initialize: (value) => { context = value; },
    onEvent(value, event) {
      if (event.type === 'mode-change') events.push({ mode: event.mode, readOnly: event.readOnly, snapshot: value.snapshot.readOnly });
    } });
  const view = render(<DocxEditor {...peerlessProps} mode="viewing" plugins={[plugin]} />);
  await waitFor(() => expect(context?.snapshot).toMatchObject({ mode: 'viewing', readOnly: true }));
  view.rerender(<DocxEditor {...peerlessProps} mode="editing" plugins={[plugin]} />);
  await waitFor(() => expect(events.at(-1)).toEqual({ mode: 'editing', readOnly: false, snapshot: false }));
});

test('suggesting shell edits carry the current host author', async () => {
  const view = render(<DocxEditor {...peerlessProps} mode="suggesting" author="Writer" />);
  await waitFor(() => expect(binding).toBeDefined());
  const first = binding!.commands.submit({ kind: 'insertText', payload: { text: 'one' } });
  if (first.status !== 'submitted') throw new Error('edit not submitted');
  await first.ticket.ack;
  view.rerender(<DocxEditor {...peerlessProps} mode="suggesting" author="Next writer" />);
  const second = binding!.commands.submit({ kind: 'insertText', payload: { text: 'two' } });
  if (second.status !== 'submitted') throw new Error('edit not submitted');
  await second.ticket.ack;
  const edits = requests.filter(request => request.type === 'edit' && request.command.kind === 'insertText');
  expect(edits.map(request => request.type === 'edit' && request.command.payload)).toMatchObject([
    { text: 'one', suggesting: { name: 'Writer' } }, { text: 'two', suggesting: { name: 'Next writer' } },
  ]);
});

test('controlled viewing mode locks admission before parent layout input', async () => {
  const submitted: { result?: ReturnType<PeerlessEditSessionBinding['commands']['submit']> } = {};
  function Parent({ mode, input = false }: { mode: 'editing' | 'viewing'; input?: boolean }) {
    useLayoutEffect(() => {
      if (input) submitted.result = binding!.commands.submit({ kind: 'insertText', payload: { text: 'locked' } });
    }, [mode, input]);
    return <DocxEditor {...peerlessProps} mode={mode} />;
  }
  const view = render(<Parent mode="editing" />);
  await waitFor(() => expect(binding).toBeDefined());
  view.rerender(<Parent mode="viewing" input />);
  expect(submitted.result).toEqual({ status: 'ignored', reason: 'read-only' });
  expect(workers[0].text).toBe('');
});


test('worker rulers show native section values and the first-line indent control', async () => {
  const view = render(<DocxEditor {...peerlessProps} showRuler />);
  await waitFor(() => expect(binding).toBeDefined());
  await waitFor(() => expect(requests.some(request => request.type === 'documentRead' && request.read.kind === 'sectionProperties')).toBe(true));
  expect(view.getByRole('slider', { name: 'First line indent' })).toBeDefined();
});

test('worker rulers use the host ruler unit', async () => {
  const view = render(<DocxEditor {...peerlessProps} showRuler rulerUnit="cm" />);
  await waitFor(() => expect(binding).toBeDefined());
  expect(view.getByRole('slider', { name: 'Horizontal ruler' }).textContent).toContain('21');
  expect(view.getByRole('slider', { name: 'Vertical ruler' }).textContent).toContain('27');
});

test('worker-only opens request the font-ready certified initial display window', async () => {
  render(<DocxEditor {...peerlessProps} showToolbar={false} />);
  await waitFor(() => expect(requests.some(request => request.type === 'commandInit')).toBe(true));
  const request = requests.find(request => request.type === 'commandInit');
  expect(request?.type === 'commandInit' && request.request).toMatchObject({
    provisionalPages: 3, displayWindow: [0, 3], layoutWithRegions: true,
  });
});


test('loading the same parsed prop document reopens exactly once for every load', async () => {
  const ref = createRef<DocxEditorRef>();
  const parsed = yrs.decodeDocxHostJson(hostJson, source).document;
  const view = render(<DocxEditor document={parsed} ref={ref} showToolbar={false} />);
  await waitFor(() => expect(ref.current!.commands.getState('save').enabled).toBe(true));
  for (const count of [2, 3]) {
    act(() => { ref.current!.loadDocument(parsed); });
    await waitFor(() => expect(workers).toHaveLength(count));
    await waitFor(() => expect(ref.current!.commands.getState('save').enabled).toBe(true));
    expect(requests.filter(({ type }) => type === 'openSource')).toHaveLength(count);
  }
  view.unmount();
});


test('each file load waits for its own bytes and opens one worker when revisiting a source', async () => {
  const ref = createRef<DocxEditorRef>();
  const file = new Blob([source]);
  const view = render(<DocxEditor documentBuffer={file} ref={ref} showToolbar={false} />);
  await waitFor(() => expect(ref.current!.commands.getState('save').enabled).toBe(true));
  let pending!: Promise<void>;
  act(() => { pending = ref.current!.loadDocumentBuffer(Uint8Array.of(8, 9)); });
  await waitFor(() => expect(workers).toHaveLength(2));
  await pending;
  let finish!: (bytes: ArrayBuffer) => void;
  const read = spyOn(file, 'arrayBuffer').mockImplementationOnce(() => new Promise<ArrayBuffer>(resolve => { finish = resolve; }));
  restorers.push(() => read.mockRestore());
  act(() => { pending = ref.current!.loadDocumentBuffer(file); });
  void pending.catch(() => {});
  try {
    await waitFor(() => expect(read).toHaveBeenCalledTimes(1));
    expect(workers).toHaveLength(2);
    act(() => { finish(new Uint8Array(source).buffer); });
    await waitFor(() => expect(workers).toHaveLength(3));
    await pending;
    expect(requests.filter(({ type }) => type === 'openSource')).toHaveLength(3);
    act(() => { pending = ref.current!.loadDocumentBuffer(file); });
    await waitFor(() => expect(workers).toHaveLength(4));
    await pending;
    expect(requests.filter(({ type }) => type === 'openSource')).toHaveLength(4);
  } finally { finish?.(new Uint8Array(source).buffer); view.unmount(); }
});


test('the real worker editor serves one copy event from its prefetched selection without the async API', async () => {
  surface.mockRestore();
  const descriptor = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
  restorers.push(() => { if (descriptor) Object.defineProperty(navigator, 'clipboard', descriptor); else Reflect.deleteProperty(navigator, 'clipboard'); });
  const onCopy = mock(() => {});
  const onError = mock(() => {});
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} onCopy={onCopy} onError={onError} />);
  const input = await view.findByRole('textbox', { name: 'Document input' });
  await act(async () => { fireEvent.keyDown(input, { key: 'a', ctrlKey: true }); });
  await waitFor(() => expect(requests.filter(r => r.type === 'documentRead' && r.read.kind === 'clipboardSlice')).toHaveLength(1));
  await act(async () => { await new Promise<void>(resolve => setImmediate(resolve)); });
  const written = new Map<string, string>();
  const data = { setData: (type: string, text: string) => { written.set(type, text); },
    getData: (type: string) => written.get(type) ?? '' } as unknown as DataTransfer;
  fireEvent.copy(input, { clipboardData: data });
  expect(data.getData('text/plain')).toBe('Selected text');
  expect(data.getData('text/html')).toBe('');
  await waitFor(() => expect(onCopy).toHaveBeenCalledTimes(1));
  expect(onError).not.toHaveBeenCalled();
});


test('unmount disposes the real worker editor clipboard subscription', async () => {
  surface.mockRestore();
  const create = clipboard.createClipboardPrefetch;
  const disposals: ReturnType<typeof mock>[] = [];
  const factory = spyOn(clipboard, 'createClipboardPrefetch').mockImplementation(options => {
    const prefetch = create(options);
    const dispose = mock(prefetch.dispose);
    disposals.push(dispose);
    return { ...prefetch, dispose };
  });
  restorers.push(() => factory.mockRestore());
  const view = render(<DocxEditor {...peerlessProps} showToolbar={false} />);
  await view.findByRole('textbox', { name: 'Document input' });
  await waitFor(() => expect(disposals).toHaveLength(1));
  view.unmount();
  expect(disposals[0]).toHaveBeenCalledTimes(1);
});
for (const readOnly of [false, true]) {
  test(`${readOnly ? 'viewer' : 'editor'} paged export uses the open layout and current render environment`, async () => {
    const ref = createRef<DocxEditorRef>();
    const element = (showHiddenText: boolean) => <DocxEditor {...peerlessProps} ref={ref} readOnly={readOnly}
      showHiddenText={showHiddenText} theme={{ colorScheme: { accent1: 'FF0000' } } as DocxEditorProps['theme']} />;
    const view = render(element(false));
    await waitFor(() => expect(view.getByTestId('peerless-paged-editor').dataset.opened).toBe('true'));
    const init = requests.find(request => request.type === 'commandInit');
    if (!init || init.type !== 'commandInit') throw new Error('Expected worker initialization');
    view.rerender(element(true));
    const options = { revisionView: 'markup' as const, stories: ['body' as const], includeGeometry: true, maxFragments: 50 };
    const result = await ref.current!.exportStructuredWithPages(options);
    expect(result.ok).toBe(true);
    if (!result.ok) throw new Error(result.failure.message);
    expect(result.content.layout.pages).toHaveLength(2);
    const exported = requests.find(request => request.type === 'documentRead' && request.read.kind === 'exportStructuredWithPages');
    if (!exported || exported.type !== 'documentRead' || exported.read.kind !== 'exportStructuredWithPages') throw new Error('Expected worker export read');
    expect(exported.read.options).toEqual(options);
    const initial = JSON.parse(init.request.layoutInput);
    expect(JSON.parse(exported.read.currentRequest)).toEqual({ ...initial, renderEnv: { ...initial.renderEnv, showHiddenText: true } });
    expect(coreSession).not.toHaveBeenCalled();
  });
}
