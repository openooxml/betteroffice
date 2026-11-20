import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { createRef } from 'react';
import * as measurement from '@betteroffice/docx/layout';
import * as yrs from '@betteroffice/docx/yrs';
import type {
  DocxCommandFrame,
  DocxCommandInitReply,
  DocxEngineWorkerHostModule,
  DocxEngineWorkerRequest,
  DocxEngineWorkerResponse,
  DocxReviewSnapshot,
} from '@betteroffice/docx/yrs';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef, type DocxErrorInfo } from '../../DocxEditor';
import { defineDocxPlugin } from '../../../plugins/defineDocxPlugin';
import * as core from '../hooks/useYrsCoreSession';
import * as sessionHook from '../hooks/usePeerlessEditSession';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { resetEngineChoiceForTests, setMissingWorkerCapabilitiesForTests } from '../internals/engineChoice';
import { loadMenuList } from '../../../../../../shared/react-toolbar/menuListLoader';

await loadMenuList();

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const source = Uint8Array.of(1, 2, 3);
const peerlessProps: DocxEditorProps = {
  documentBuffer: source,
};
const requests: DocxEngineWorkerRequest[] = [];
const workers: FakeWorker[] = [];
let binding: PeerlessEditSessionBinding | undefined;

const hostJson = JSON.stringify({
  referencedFonts: [] as string[],
  envelope: {
    wireVersion: 1,
    document: {
      package: {
        document: { content: [], sections: [] },
        theme: {}, numbering: {}, settings: {}, fontTable: { fonts: [] },
        relationshipEntries: [], mediaEntries: [], chartEntries: [],
      },
    },
    embeddedFontParts: [],
  },
});

function commandFrame(epoch: number): DocxCommandFrame {
  const bytes = new Uint8Array(88);
  bytes.set([0x46, 0x44, 0x56, 0x31]);
  const data = new DataView(bytes.buffer);
  data.setUint16(4, 1, true);
  data.setUint16(6, 80, true);
  data.setUint32(8, bytes.length, true);
  data.setUint32(12, 1, true);
  data.setBigUint64(16, 1n, true);
  data.setBigUint64(24, 1n, true);
  data.setBigUint64(32, BigInt(epoch), true);
  data.setBigUint64(40, 0n, true);
  data.setUint32(56, 80, true);
  data.setUint32(60, 80, true);
  data.setUint32(64, 4, true);
  data.setUint32(68, 88, true);
  return {
    coversThroughSeq: epoch - 1, revision: epoch,
    payload: { frame: bytes.buffer, caret: { frameEpoch: epoch, caretRect: null }, layoutRevision: epoch,
      stagedPresentation: { frameEpoch: epoch, pages: [] } },
  };
}

function reviewSnapshot(version: string, frameEpoch: number): DocxReviewSnapshot {
  return {
    version, previewVersion: 0, frameEpoch, layoutRevision: frameEpoch,
    proposals: { version, previewVersion: 0, proposals: [] },
    proposalGeometry: { version, previewVersion: 0, proposals: '', targets: {}, hidden: [] },
    proposalHighlights: {}, comments: [], uiComments: [], commentIds: {}, revisions: [], commentGeometry: {}, revisionGeometry: {},
  };
}

class FakeWorker {
  onmessage: ((event: MessageEvent<DocxEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  terminate = mock(() => {});
  initial: DocxCommandInitReply | undefined;
  text = '';
  frameEpoch = 0;
  fontRevision = 0;

  constructor() {
    workers.push(this);
  }

  documentVersion() {
    return `worker-${workers.indexOf(this) + 1}-${this.text.length}`;
  }

  frame(throughSeq: number): DocxCommandFrame {
    this.frameEpoch = throughSeq + 1;
    const frame = commandFrame(this.frameEpoch);
    return { ...frame, coversThroughSeq: throughSeq, payload: { ...frame.payload, documentVersion: this.documentVersion() } };
  }

  postMessage(message: DocxEngineWorkerRequest | DocxEngineWorkerHostModule) {
    if (!('id' in message)) return;
    requests.push(message);
    if (message.type === 'prepare-save') return;
    let response: DocxEngineWorkerResponse;
    if (message.type === 'openSource') {
      this.text = message.source.state ? new TextDecoder().decode(message.source.state) : '';
      response = { id: message.id, ok: true, hostJson, sourceDocument: yrs.decodeDocxHostJson(hostJson, source).document, stateVector: new ArrayBuffer(0) };
    } else if (message.type === 'fontRequirements') {
      response = { id: message.id, ok: true, requirementsJson: '[]' };
    } else if (message.type === 'save') {
      response = { id: message.id, ok: true, saved: new TextEncoder().encode(this.text).buffer, updates: [], version: 'saved' };
    } else if (message.type === 'documentRead' && message.read.kind === 'reviewSnapshot') {
      const version = this.documentVersion();
      response = !this.initial ? { id: message.id, ok: false, error: 'Document worker is not initialized' }
        : message.expectVersion !== undefined && message.expectVersion !== version ? { id: message.id, ok: true, superseded: true }
        : { id: message.id, ok: true, read: { version, value: reviewSnapshot(version, this.frameEpoch) } };
    } else if (message.type === 'documentRead') {
      response = { id: message.id, ok: true, read: { version: 'worker-version', value: message.read.kind === 'clipboardSlice' ? { text: 'Selected text', fragment: null }
        : message.read.kind === 'headings'
        ? [{ story: 'body', paraId: 'heading', text: 'Worker heading', level: 0, position: 1 }]
        : message.read.kind === 'sectionProperties' ? { version: 'worker-version', sections: [{ section: 'final', properties: { marginLeft: 2160 } }], paragraphs: [] }
        : message.read.kind === 'searchText' ? { matches: [], carried: -1 }
        : message.read.kind === 'rangeHighlights' || message.read.kind === 'sidebar' ? null
        : message.read.kind === 'toolbarMetadata' ? { styles: [], fontFamilies: [] }
        : message.read.kind === 'listStories' ? { ok: true, version: 'worker-version', stories: [{ story: 'hf:rId3', kind: 'header', root: 'hf:rId3' }] }
        : message.read.kind === 'readStories' ? { ok: true, version: 'worker-version', view: 'accepted', stories: [{ story: 'hf:rId3', ok: true, paragraphs: [] }], truncated: false }
        : { ok: true, view: 'accepted', version: 'worker-version', paragraphs: [] } } };
    } else if (message.type === 'requestFullFrame' && this.initial) {
      response = { id: message.id, ok: true, commandAck: { ...this.initial, frame: this.frame(this.initial.throughSeq) } };
    } else if (message.type === 'commandInit') {
      const { session, incarnation, epoch, version = 0, throughSeq = 0, readOnly } = message.request;
      const documentVersion = this.documentVersion();
      const initial: DocxCommandInitReply = {
        session, incarnation, epoch, version, throughSeq, outcomes: [], selection: null,
        editState: {
          readOnly, toolbar: null, effectiveFormatting: null, storedMarks: null, documentVersion,
          review: { version: documentVersion, previewVersion: 0 },
          canUndo: false, canRedo: false, hyperlink: null, tableContext: null,
        },
        frame: this.frame(throughSeq),
      };
      this.initial = initial;
      response = { id: message.id, ok: true, commandInitial: initial };
    } else if (message.type === 'edit' && this.initial) {
      const command = message.command;
      const state = this.initial;
      const readOnly = command.kind === 'setReadOnly' ? (command.payload as { readOnly: boolean }).readOnly : state.editState.readOnly;
      if (command.kind === 'insertText') this.text += (command.payload as { text: string }).text;
      const documentVersion = this.documentVersion();
      const ack = {
        session: state.session, incarnation: state.incarnation, epoch: state.epoch,
        version: state.version + (command.kind === 'insertText' ? 1 : 0), throughSeq: command.seq,
        outcomes: [{ seq: command.seq, status: command.kind === 'insertText' ? 'applied' as const : 'noop' as const }],
        editState: { ...state.editState, readOnly, documentVersion, review: { version: documentVersion, previewVersion: 0 } },
        frame: this.frame(command.seq),
      };
      this.initial = { ...state, ...ack };
      response = { id: message.id, ok: true, commandAck: ack };
      if (command.kind === 'setFontEnvironment') {
        response.fontEnvironmentEvent = { session: state.session, incarnation: state.incarnation, epoch: state.epoch,
          phase: 'ready', revision: ++this.fontRevision, families: (command.payload as yrs.DocxCommandPayloads['setFontEnvironment']).families };
      }
    } else {
      response = { id: message.id, ok: true };
    }
    queueMicrotask(() => this.onmessage?.({ data: response } as MessageEvent<DocxEngineWorkerResponse>));
  }
}

beforeEach(() => {
  requests.length = 0;
  workers.length = 0;
  binding = undefined;
  setMissingWorkerCapabilitiesForTests([]);
  const canvas = window.HTMLCanvasElement.prototype as { transferControlToOffscreen?: () => unknown };
  const transfer = Object.getOwnPropertyDescriptor(canvas, 'transferControlToOffscreen');
  Object.defineProperty(canvas, 'transferControlToOffscreen', { configurable: true, value: () => ({}) });
  restorers.push(() => {
    if (transfer) Object.defineProperty(canvas, 'transferControlToOffscreen', transfer);
    else Reflect.deleteProperty(canvas, 'transferControlToOffscreen');
  });
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
  const coreSession = spyOn(core, 'useYrsCoreSession').mockReturnValue({
    session: null, sessionGeneration: null, replicaReady: false, workerProposalsReady: false,
    requestReplica() {}, storyBlocks: () => null, bodyBlocks: () => null,
    inputPositionMap: () => null, displayPositionToLoc: () => null, locToDisplayPosition: () => null,
    documentFromYrs: () => null, publishDirectInput() {}, previewing: false, handoffFrom: null,
    opening: false, notifyFramePresented() {}, failOpening: () => false,
    scheduleCompatibilityWarm() {}, cancelCompatibilityWarm() {},
  });
  restorers.push(() => coreSession.mockRestore());
  const original = sessionHook.usePeerlessEditSession;
  const hook = spyOn(sessionHook, 'usePeerlessEditSession').mockImplementation(options => {
    const result = original(options);
    if (result.binding) binding = result.binding;
    return result;
  });
  restorers.push(() => hook.mockRestore());
});

afterEach(async () => {
  cleanup();
  await act(async () => { await new Promise<void>(resolve => setTimeout(resolve, 0)); });
  for (const restore of restorers.splice(0).reverse()) restore();
  resetEngineChoiceForTests();
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('with paste options off and a plugin, a worker being reopened changes no view, reads nothing and reports nothing to the host', async () => {
  const proposalVersions: string[] = [];
  const plugin = defineDocxPlugin({ id: 'recovering-watch', createState: () => null,
    async onEvent(_context, event) {
      if (event.type === 'proposal-change') proposalVersions.push(event.version);
    },
  });
  const ref = createRef<DocxEditorRef>();
  const onError = mock((_error: Error, _info?: DocxErrorInfo) => {});
  render(<DocxEditor {...peerlessProps} ref={ref} pasteOptions={false} plugins={[plugin]} onError={onError} showToolbar={false} />);
  await waitFor(() => expect(binding).toBeDefined());
  await act(async () => {
    const submitted = binding!.commands.submit({ kind: 'insertText', payload: { text: 'kept' } });
    if (submitted.status !== 'submitted') throw new Error('edit not submitted');
    expect((await submitted.ticket.ack).status).toBe('applied');
  });
  const kept = binding!.session.state.throughSeq;
  await waitFor(() => expect(binding!.session.state.presentedThroughSeq).toBe(kept));
  await waitFor(() => expect(proposalVersions).toContain('worker-1-4'));
  await act(async () => { await new Promise<void>(resolve => setImmediate(resolve)); });
  const viewStore = binding!.viewStore;
  const view = viewStore.getSnapshot();
  let notified = 0;
  const stop = viewStore.subscribe(() => { notified += 1; });
  const reads = requests.filter(request => request.type === 'documentRead').length;
  await act(async () => { viewStore.setRecovering?.(true); });
  await act(async () => { await new Promise<void>(resolve => setImmediate(resolve)); });
  stop();
  expect(viewStore.getRecovering?.()).toBe(true);
  expect(notified).toBe(0);
  expect(viewStore.getSnapshot()).toBe(view);
  expect(requests.filter(request => request.type === 'documentRead').length).toBe(reads);
  expect(onError).not.toHaveBeenCalled();
});
