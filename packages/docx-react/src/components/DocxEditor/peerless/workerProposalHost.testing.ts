import { mock, type Mock } from 'bun:test';
import type { DocxProposalInput, DocxProposalSnapshot, DocxWorkerEditSession, DocxReviewSnapshot } from '@betteroffice/docx/yrs';
import type { DocxEditorProps } from '../../DocxEditor';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import { createInputAdmission } from './inputAdmission';
import { createPeerlessCommandBinding } from './commandBinding';
import { createPeerlessShellController } from './shellController';
import { createPeerlessViewStore, type PeerlessViewSnapshot } from './viewStore';

type ProposalTicket = ReturnType<DocxWorkerEditSession<PeerlessEchoIntent>['submit']>;
type ReviewRead = { status: 'superseded' } | { status: 'ok'; version: string; value: DocxReviewSnapshot | null };
interface ProposalHost {
  controller: ReturnType<typeof createPeerlessShellController>;
  binding: PeerlessEditSessionBinding;
  submit: Mock<(kind: string, payload: unknown, metadata: unknown) => ProposalTicket>;
  readAt: Mock<(request: { kind: string; expectPreviewVersion?: number }, version: string) => Promise<ReviewRead>>;
  present(): PeerlessViewSnapshot;
  settings: Parameters<typeof createPeerlessShellController>[0];
  snapshot(): DocxProposalSnapshot;
  nullReads(): number;
  refuse(reason: string): void;
  assignIdentities(): void;
}

export const source = Uint8Array.of(1);

export const inputs: DocxProposalInput[] = Array.from({ length: 4 }, (_, index) => ({
  id: `change-${index}`, paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: `p${index}` },
  op: 'insertText', at: 'end', text: `Added ${index}`, suggest: { author: 'Reviewer', date: '2026-10-01T00:00:00Z' },
}));

export function proposalHost(options: { readOnly?: boolean; mode?: DocxEditorProps['mode']; allowHostProposals?: boolean;
  automatic?: boolean; opening?: boolean } = {}): ProposalHost {
  const workerReadOnly = options.readOnly === true || options.mode === 'viewing';
  const identity = { session: 's', incarnation: 1, epoch: 1 };
  let snapshot: DocxProposalSnapshot = { version: 'native-0', previewVersion: 0, proposals: [] };
  let seq = 0;
  let frameEpoch = 0;
  let refuse: string | undefined;
  let nullReads = 0;
  const viewStore = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0,
    workerSelection: null, anchors: null, geometry: null, editState: {
      documentVersion: snapshot.version, review: { version: snapshot.version, previewVersion: 0 }, readOnly: workerReadOnly,
      toolbar: null, effectiveFormatting: null, storedMarks: null, canUndo: true, canRedo: false, hyperlink: null, tableContext: null,
    } });
  const present = () => {
    const view = viewStore.getSnapshot();
    viewStore.installFrameGeometry({ identity, documentVersion: snapshot.version, coversThroughSeq: seq,
      frameEpoch: ++frameEpoch, layoutRevision: frameEpoch, pages: [], caret: null, selectionRects: [] });
    return view;
  };
  present();
  const submit = mock((kind: string, payload: unknown, _metadata: unknown): ProposalTicket => {
    const request = payload as { proposals: DocxProposalInput[]; changes: { id: string; state: 'proposed' | 'accepted' | 'rejected' }[]; ids: string[]; hostProposal?: boolean };
    const currentSeq = ++seq;
    const permitted = request.hostProposal === true || ['setReadOnly', 'setRenderEnvironment', 'setFontEnvironment'].includes(kind);
    const reason = refuse ?? (workerReadOnly && !permitted ? 'read-only' : undefined);
    refuse = undefined;
    const ack = Promise.resolve().then(() => {
      if (!reason) {
        if (kind === 'propose') snapshot = { ...snapshot, version: `native-${seq}`, proposals: request.proposals.map(input => ({
          id: input.id, paragraph: input.paragraph as never, state: 'proposed', revisionIds: [`revision-${input.id}`], changed: true,
        })) };
        if (kind === 'setProposalStates') snapshot = { ...snapshot, previewVersion: snapshot.previewVersion + 1,
          proposals: snapshot.proposals.map(record => ({ ...record, state: request.changes.find(change => change.id === record.id)?.state ?? record.state })) };
        if (kind === 'withdrawProposals') snapshot = { ...snapshot, version: `native-${seq}`, proposals: snapshot.proposals.filter(record => !request.ids.includes(record.id)) };
        if (kind === 'undo' || kind === 'redo') snapshot = { ...snapshot, version: `native-${seq}` };
      }
      const view = viewStore.getSnapshot();
      viewStore.acceptWorkerState(identity, { epoch: 1, version: view.version + (reason ? 0 : 1), throughSeq: currentSeq,
        presentedThroughSeq: view.presentedThroughSeq, selection: null, editState: { ...view.editState!, documentVersion: snapshot.version,
          review: { version: snapshot.version, previewVersion: snapshot.previewVersion } } });
      if (options.automatic !== false) present();
      return { status: reason ? 'refused' as const : 'applied' as const, seq: currentSeq, ...(reason ? { reason } : {}) };
    });
    return { id: currentSeq, seq: currentSeq, ack, presented: Promise.resolve('presented' as const) };
  });
  const session = { get state() { const view = viewStore.getSnapshot(); return { ...view, epoch: 1, selection: view.workerSelection }; },
    submit, flush: mock(async () => {}) } as unknown as DocxWorkerEditSession<PeerlessEchoIntent>;
  const admission = createInputAdmission({ session, viewStore, echoPainter: { disable() {}, preview: () => undefined } as unknown as PeerlessEchoPainter });
  const binding = { session, viewStore, admission, commands: createPeerlessCommandBinding(admission) } as PeerlessEditSessionBinding;
  const readAt = mock(async (request: { kind: string; expectPreviewVersion?: number }, version: string): Promise<ReviewRead> => {
    const frame = viewStore.getSnapshot().geometry!;
    if (version !== snapshot.version) return { status: 'superseded' as const };
    const available = frame.documentVersion === snapshot.version && frame.coversThroughSeq >= seq && request.expectPreviewVersion === snapshot.previewVersion;
    if (!available) nullReads++;
    const value: DocxReviewSnapshot | null = available ? {
      version: snapshot.version, previewVersion: snapshot.previewVersion, frameEpoch: frame.frameEpoch, layoutRevision: frame.layoutRevision,
      proposals: structuredClone(snapshot), proposalGeometry: { version: snapshot.version, previewVersion: snapshot.previewVersion,
        proposals: '', targets: Object.fromEntries(snapshot.proposals.map((proposal, index) => [proposal.id, {
          ok: true, ranges: [{ from: index * 10, to: index * 10 + 5 }], paragraph: index * 10,
        }])), hidden: [] }, proposalHighlights: Object.fromEntries(snapshot.proposals.map((proposal, index) => [proposal.id,
          [{ pageIndex: index < 2 ? 0 : 1, x: 30, y: 40 + (index % 2) * 20, width: 50, height: 12 }]])),
      comments: [], uiComments: [], commentIds: {}, revisions: snapshot.proposals.map((proposal, index) => ({
        revisionId: proposal.revisionIds[0]!, kind: 'insertion', author: 'Reviewer', date: '2026-10-01T00:00:00Z', preview: `Added ${index}`,
        range: { story: 'body', start: { paraId: `p${index}`, offset: 0 }, end: { paraId: `p${index}`, offset: 5 } },
      })) as DocxReviewSnapshot['revisions'], commentGeometry: {}, revisionGeometry: {},
    } : null;
    return { status: 'ok' as const, version, value };
  });
  const read = mock(async () => ({ version: snapshot.version, value: { ok: true, view: 'accepted', version: snapshot.version, paragraphs: [] } }));
  const settings = { read: read as never, readAt: readAt as never, focus() {}, zoom: 1, readOnly: options.readOnly, mode: options.mode,
    allowHostProposals: options.allowHostProposals ?? true };
  const controller = createPeerlessShellController(settings);
  if (options.opening) controller.opening(source);
  else controller.attach(binding);
  return { controller, binding, submit, readAt, present, settings, snapshot: () => snapshot, nullReads: () => nullReads,
    refuse: (reason: string) => { refuse = reason; }, assignIdentities: () => { snapshot = { ...snapshot, version: `${snapshot.version}+ids` }; } };
}
